App({
  onLaunch: function () {
    if (!wx.cloud) {
      console.error('请使用 2.2.3 或以上的基础库以使用云能力')
    } else {
      wx.cloud.init({
        env: 'your-env-id',
        traceUser: true,
      })
    }

    this.globalData = {
      userInfo: null,
      openid: null
    }
  },

  setUserInfo: function (userInfo, openid) {
    this.globalData.userInfo = userInfo
    this.globalData.openid = openid
  }
})