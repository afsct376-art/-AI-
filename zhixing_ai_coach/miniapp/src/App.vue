<script>
const ENV_ID = 'zhixing-d1gp1p0aj9744d967'

export default {
  onLaunch() {
    // 初始化 CloudBase
    if (wx.cloud) {
      wx.cloud.init({ env: ENV_ID, traceUser: true })
    }
  }
}
</script>

<style>
/* #ifndef APP-PLUS-NVUE */
page {
  font-family: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
  background-color: #f5f5f5;
  color: #333;
  font-size: 28rpx;
  line-height: 1.6;
}
/* #endif */
</style>
