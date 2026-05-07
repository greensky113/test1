const app = getApp()

Page({
  data: {
    loading: false
  },

  onLoad: function (options) {
    this.checkLoginStatus()
  },

  checkLoginStatus: function () {
    const openid = wx.getStorageSync('openid')
    const userInfo = wx.getStorageSync('userInfo')
    if (openid && userInfo) {
      app.setUserInfo(userInfo, openid)
      wx.switchTab({
        url: '/pages/index/index'
      })
    }
  },

  handleLogin: function (e) {
    if (e.detail.userInfo) {
      this.setData({ loading: true })
      
      wx.cloud.callFunction({
        name: 'login',
        data: {},
        success: res => {
          const openid = res.result.openid
          const userInfo = e.detail.userInfo
          
          wx.setStorageSync('openid', openid)
          wx.setStorageSync('userInfo', userInfo)
          app.setUserInfo(userInfo, openid)
          
          wx.showToast({
            title: '登录成功',
            icon: 'success',
            duration: 1500
          })
          
          setTimeout(() => {
            this.setData({ loading: false })
            wx.switchTab({
              url: '/pages/index/index'
            })
          }, 1500)
        },
        fail: err => {
          this.setData({ loading: false })
          wx.showToast({
            title: '登录失败',
            icon: 'none'
          })
        }
      })
    } else {
      wx.showToast({
        title: '需要授权才能使用',
        icon: 'none'
      })
    }
  }
})