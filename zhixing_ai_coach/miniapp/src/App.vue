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
  --primary-color: #e94560;
  --dark-color: #1a1a2e;
  --bg-color: #f7f8fa;
  --radius-lg: 20rpx;
  --radius-md: 12rpx;
  --shadow-base: 0 4rpx 12rpx rgba(0,0,0,0.05);
  
  font-family: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif;
  background-color: var(--bg-color);
  color: #333;
  font-size: 28rpx;
  line-height: 1.6;
}
/* #endif */
</style>
