// #ifndef VUE3
import Vue from 'vue'
import Vuex from 'vuex'
Vue.use(Vuex)

const store = new Vuex.Store({
  state: {
    token: uni.getStorageSync('token') || '',
    user: null,
    currentMember: null,
  },
  mutations: {
    SET_TOKEN(state, token) {
      state.token = token
      uni.setStorageSync('token', token)
    },
    SET_USER(state, user) {
      state.user = user
    },
    SET_CURRENT_MEMBER(state, member) {
      state.currentMember = member
    },
    CLEAR_AUTH(state) {
      state.token = ''
      state.user = null
      uni.removeStorageSync('token')
    },
  },
  actions: {
    async login({ commit }, code) {
      const res = await uni.request({
        url: `${BASE_URL}/api/auth/login`,
        method: 'POST',
        data: { code },
      })
      if (res.data.access_token) {
        commit('SET_TOKEN', res.data.access_token)
        commit('SET_USER', res.data.user)
      }
      return res.data
    },
  },
  getters: {
    isLoggedIn: state => !!state.token,
  },
})

export default store
// #endif

// #ifdef VUE3
import { createStore } from 'vuex'

const store = createStore({
  state: {
    token: uni.getStorageSync('token') || '',
    user: null,
    currentMember: null,
  },
  mutations: {
    SET_TOKEN(state, token) {
      state.token = token
      uni.setStorageSync('token', token)
    },
    SET_USER(state, user) {
      state.user = user
    },
    SET_CURRENT_MEMBER(state, member) {
      state.currentMember = member
    },
    CLEAR_AUTH(state) {
      state.token = ''
      state.user = null
      uni.removeStorageSync('token')
    },
  },
  actions: {
    async login({ commit }, code) {
      const res = await uni.request({
        url: `${BASE_URL}/api/auth/login`,
        method: 'POST',
        data: { code },
      })
      if (res.data.access_token) {
        commit('SET_TOKEN', res.data.access_token)
        commit('SET_USER', res.data.user)
      }
      return res.data
    },
  },
  getters: {
    isLoggedIn: state => !!state.token,
  },
})

export default store
// #endif
