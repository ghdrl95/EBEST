import tensorflow as tf
import numpy as np
import random
from collections import deque
#학습모델
class DQN:
    REPLAY_MEMORY=10000 # 학습에 사용할 플레이결과를 얼마나 저장할지
    BATCH_SIZE=32 #한번 학습시 몇개의 기억을 사용할지 정하는것
    GAMMA = 0.99 # 학습가중치
    STATE_LEN = 4 #한번에 볼 프레임 수
    def __init__(self,session,n_action):
        self.session = session
        self.n_action = n_action
        self.memory = deque() ## 게임 플레이 결과를 저장할 메모리를 만드는 코드
        #parameter settings
        self.input_X=tf.placeholder(tf.float32 , [None,7,150,self.STATE_LEN]) #게임 상태
        self.input_A=tf.placeholder(tf.int64 , [None])                                #각 상태를 만들어낸 액션의 값
        self.input_Y=tf.placeholder(tf.float32 , [None])                              #손실값 계산에 사용할 변수

        self.Q = self._build_network('main')
        self.cost, self.train_op = self._build_op()

        self.target_Q = self._build_network('target')

    #학습 신경망과 목표 신경망을 구성하는 함수, 상탯값 X를 받아 행동의 가짓수만큼의 출력값을 만듬, 최대값을 취해 다음행동 결정
    def _build_network(self,name):
        with tf.variable_scope(name):
            model = tf.layers.conv2d(self.input_X,32,[4,4], padding='same', activation=tf.nn.relu)
            model = tf.layers.conv2d(model, 65, [2,2], padding='same', activation=tf.nn.relu)
            model = tf.contrib.layers.flatten(model)
            model = tf.layers.dense(model,512,activation=tf.nn.relu)

            Q = tf.layers.dense(model,self.n_action,activation=None)
        return Q
    # DQN의 손실 함수를 구하는 부분
    def _build_op(self):
        one_hot = tf.one_hot(self.input_A,self.n_action, 1.0, 0.0) #활성화된 액션에 1.0 아닌거에 0.0을 입력한 리스트로 세팅하는거
        #tf.multiply(self.Q,one_hot) 현재 행동의 인덱스만 해당하는 값을 쓰기위해서 사용
        Q_value = tf.reduce_sum(tf.multiply(self.Q,one_hot), axis=1) #리스트에 있는거 다 합치는거
        cost = tf.reduce_mean(tf.square(self.input_Y-Q_value))
        train_op = tf.train.AdamOptimizer(1e-6).minimize(cost)

        return cost, train_op

    #목표 신경망 갱신
    def update_target_network(self):
        copy_op =[]
        main_vars = tf.get_collection(tf.GraphKeys.TRAINABLE_VARIABLES,scope='main')
        target_vars = tf.get_collection(tf.GraphKeys.TRAINABLE_VARIABLES,scope='target')

        for main_var, target_var in zip(main_vars, target_vars):
            copy_op.append(target_var.assign(main_var.value()))

        self.session.run(copy_op)

    #현재 상태를 이용해 다음에 취할 행동을 찾는 함수
    def get_action(self):
        Q_value = self.session.run(self.Q,feed_dict={self.input_X: [self.state]})
        action = np.argmax(Q_value[0])
        return action

    #학습
    def train(self):
        state,next_state,action,reward,terminal = self._sample_memory()
        target_Q_value = self.session.run(self.target_Q,feed_dict={self.input_X:next_state})
        Y=[]
        for i in range(self.BATCH_SIZE):
            if terminal[i]:
                Y.append(reward[i])
            else:
                Y.append(reward[i]+self.GAMMA * np.max(target_Q_value[i]))
        self.session.run(self.train_op,feed_dict={self.input_X:state, self.input_A:action, self.input_Y:Y})
    #상태 초기화
    def init_state(self,state):
        state = [state for _ in range(self.STATE_LEN)]
        self.state = np.stack(state,axis=2)


    #게임 결과를 받아 메모리에 기억하는 기능
    def remember(self,state,action,reward, terminal):
        next_state = np.reshape(state,(self.width,self.height,1))
        next_state = np.append(self.state[:,:,1:],next_state,axis=2)

        self.memory.append((self.state,next_state,action,reward,terminal))

        if (len(self.memory)>self.REPLAY_MEMORY):
            self.memory.popleft()
        self.state=next_state

    def _sample_memory(self):
        sample_memory=random.sample(self.memory, self.BATCH_SIZE)

        state = [memory[0] for memory in sample_memory]
        next_state = [memory[1] for memory in sample_memory]
        action = [memory[2] for memory in sample_memory]
        reward = [memory[3] for memory in sample_memory]
        terminal = [memory[4] for memory in sample_memory]

        return state, next_state, action, reward, terminal