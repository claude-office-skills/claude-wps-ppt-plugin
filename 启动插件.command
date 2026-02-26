#!/bin/bash
cd "$(dirname "$0")"

echo "==================================="
echo " WPS Claude PPT 插件启动器"
echo " 端口: 3002 (proxy) / 5174 (vite)"
echo " ※ 不影响 Excel 插件 (3001/5173)"
echo "==================================="

# 检查 node_modules
if [ ! -d "node_modules" ]; then
  echo "📦 正在安装依赖..."
  npm install
fi

# 检查 claude CLI
if ! command -v claude &> /dev/null; then
  echo "❌ 未找到 claude CLI，请先安装 Claude Code"
  read -p "按任意键退出..." k
  exit 1
fi

# 停止已有 PPT 插件实例（仅清理 PPT 端口，不影响 Excel）
echo "🛑 清理 PPT 插件旧进程..."
lsof -ti:3002 | xargs kill -9 2>/dev/null || true
lsof -ti:5174 | xargs kill -9 2>/dev/null || true
sleep 1

echo ""
echo "▶ 启动代理服务器 (端口 3002)..."
node proxy-server.js > /tmp/proxy-server-ppt.log 2>&1 &
PROXY_PID=$!
sleep 2

if curl -s http://127.0.0.1:3002/health > /dev/null 2>&1; then
  echo "   ✅ 代理服务器启动成功"
else
  echo "   ❌ 代理服务器启动失败"
  cat /tmp/proxy-server-ppt.log
fi

echo "▶ 启动前端开发服务器 (端口 5174)..."
npm run dev > /tmp/vite-dev-ppt.log 2>&1 &
VITE_PID=$!

echo "   等待 Vite 就绪..."
for i in $(seq 1 15); do
  sleep 1
  if curl -s http://127.0.0.1:5174/ > /dev/null 2>&1; then
    echo "   ✅ 前端服务器就绪 (${i}s)"
    break
  fi
  echo -n "."
done

echo ""
echo "==================================="
echo " ✅ PPT 插件服务已启动！"
echo ""
echo " 前端: http://127.0.0.1:5174"
echo " 代理: http://127.0.0.1:3002/health"
echo ""
echo " 请在 WPS PPT 中点击 [Claude AI] → [Claude 助手]"
echo " 如未出现按钮，请运行: ./install-to-wps.sh"
echo ""
echo " ⚠️  关闭此窗口将停止 PPT 插件服务！"
echo " ⚠️  Excel 插件 (3001/5173) 不受影响。"
echo "==================================="

trap "echo ''; echo '关闭 PPT 插件服务...'; kill $PROXY_PID $VITE_PID 2>/dev/null; exit" INT TERM
wait
