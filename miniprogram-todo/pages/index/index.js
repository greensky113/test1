const app = getApp()
const db = wx.cloud.database()

Page({
  data: {
    userInfo: null,
    todos: [],
    inputValue: '',
    currentFilter: 'all',
    loading: false
  },

  onLoad: function (options) {
    this.checkLogin()
  },

  onShow: function () {
    this.loadTodos()
  },

  checkLogin: function () {
    const openid = wx.getStorageSync('openid')
    const userInfo = wx.getStorageSync('userInfo')
    
    if (!openid || !userInfo) {
      wx.redirectTo({
        url: '/pages/login/login'
      })
      return
    }
    
    app.setUserInfo(userInfo, openid)
    this.setData({ userInfo })
  },

  loadTodos: function () {
    const openid = wx.getStorageSync('openid')
    if (!openid) return

    this.setData({ loading: true })
    
    db.collection('todos')
      .where({
        _openid: openid
      })
      .orderBy('createTime', 'desc')
      .get({
        success: res => {
          this.setData({
            todos: res.data,
            loading: false
          })
        },
        fail: err => {
          console.error('加载失败', err)
          this.setData({ loading: false })
        }
      })
  },

  handleInput: function (e) {
    this.setData({
      inputValue: e.detail.value
    })
  },

  addTodo: function () {
    const content = this.data.inputValue.trim()
    if (!content) return

    const openid = wx.getStorageSync('openid')
    if (!openid) return

    const now = new Date()
    const createTime = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`

    db.collection('todos').add({
      data: {
        content,
        completed: false,
        createTime
      },
      success: res => {
        wx.showToast({
          title: '添加成功',
          icon: 'success',
          duration: 1000
        })
        
        this.setData({ inputValue: '' })
        this.loadTodos()
      },
      fail: err => {
        wx.showToast({
          title: '添加失败',
          icon: 'none'
        })
      }
    })
  },

  toggleTodo: function (e) {
    const id = e.currentTarget.dataset.id
    const todo = this.data.todos.find(t => t._id === id)
    
    db.collection('todos').doc(id).update({
      data: {
        completed: !todo.completed
      },
      success: res => {
        this.loadTodos()
      },
      fail: err => {
        wx.showToast({
          title: '操作失败',
          icon: 'none'
        })
      }
    })
  },

  deleteTodo: function (e) {
    const id = e.currentTarget.dataset.id
    
    wx.showModal({
      title: '确认删除',
      content: '确定要删除这个任务吗？',
      success: res => {
        if (res.confirm) {
          db.collection('todos').doc(id).remove({
            success: () => {
              wx.showToast({
                title: '删除成功',
                icon: 'success',
                duration: 1000
              })
              this.loadTodos()
            },
            fail: () => {
              wx.showToast({
                title: '删除失败',
                icon: 'none'
              })
            }
          })
        }
      }
    })
  },

  setFilter: function (e) {
    const filter = e.currentTarget.dataset.filter
    this.setData({ currentFilter: filter })
  },

  get filteredTodos() {
    const { todos, currentFilter } = this.data
    switch (currentFilter) {
      case 'pending':
        return todos.filter(todo => !todo.completed)
      case 'completed':
        return todos.filter(todo => todo.completed)
      default:
        return todos
    }
  },

  get totalCount() {
    return this.data.todos.length
  },

  get completedCount() {
    return this.data.todos.filter(todo => todo.completed).length
  },

  get pendingCount() {
    return this.data.todos.filter(todo => !todo.completed).length
  },

  get emptyText() {
    const { currentFilter } = this.data
    switch (currentFilter) {
      case 'pending':
        return '太棒了！没有待完成的任务'
      case 'completed':
        return '还没有已完成的任务'
      default:
        return '暂无任务，快来添加一个吧'
    }
  }
})