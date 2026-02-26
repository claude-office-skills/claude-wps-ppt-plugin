#!/bin/bash
# 安装 Claude PPT 插件到 WPS（直接写入 jsaddons 配置）
# PPT 插件使用端口 3002/5174，与 Excel 插件 (3001/5173) 完全独立

echo "==================================="
echo " Claude for WPS PPT 插件安装脚本"
echo " 端口: 3002 (proxy) / 5174 (vite)"
echo "==================================="

cd "$(dirname "$0")"

WPS_JSADDONS="$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons"

# 1. 启动代理服务器（仅清理 PPT 端口）
echo ""
echo "▶ [1/4] 启动代理服务器 (端口 3002)..."
lsof -ti:3002 | xargs kill -9 2>/dev/null
sleep 1
node proxy-server.js &
PROXY_PID=$!
sleep 2

if curl -s http://127.0.0.1:3002/health > /dev/null 2>&1; then
  echo "   ✅ 代理服务器启动成功"
else
  echo "   ❌ 代理服务器启动失败"
fi

# 2. 启动 Vite 前端（仅清理 PPT 端口）
echo "▶ [2/4] 启动前端服务器 (端口 5174)..."
lsof -ti:5174 | xargs kill -9 2>/dev/null
sleep 1
npm run dev &
VITE_PID=$!
sleep 4

if curl -s http://127.0.0.1:5174/ > /dev/null 2>&1; then
  echo "   ✅ 前端服务器启动成功"
else
  echo "   ⚠️  前端服务器可能还在启动中..."
fi

# 3. 写入 WPS jsaddons 配置
echo "▶ [3/4] 注册 PPT 插件到 WPS..."

if [ ! -d "$WPS_JSADDONS" ]; then
  echo "   ❌ WPS jsaddons 目录不存在: $WPS_JSADDONS"
  echo "   请先安装并启动过 WPS Office"
  kill $PROXY_PID $VITE_PID 2>/dev/null
  exit 1
fi

PUBLISH_XML="$WPS_JSADDONS/publish.xml"
PPT_ENTRY='<jspluginonline name="claude-wps-ppt-plugin" type="wpp" url="http://127.0.0.1:3002/wps-addon/" debug="" enable="enable_dev" install="null"/>'

if grep -q "claude-wps-ppt-plugin" "$PUBLISH_XML" 2>/dev/null; then
  echo "   ✅ publish.xml 已包含 PPT 插件条目"
else
  sed -i '' "s|</jsplugins>|  $PPT_ENTRY\n</jsplugins>|" "$PUBLISH_XML"
  echo "   ✅ 已添加 PPT 插件到 publish.xml"
fi

python3 -c "
import json, sys
path = '$WPS_JSADDONS/authaddin.json'
try:
    with open(path) as f:
        data = json.load(f)
except:
    data = {}
data['wpp'] = {
    '70707450314b49546944497842773274': {
        'enable': True,
        'isload': True,
        'md5': '',
        'mode': 2,
        'name': 'claude-wps-ppt-plugin',
        'path': 'http://127.0.0.1:3002/wps-addon'
    },
    'namelist': '70707450314b49546944497842773274'
}
with open(path, 'w') as f:
    json.dump(data, f, indent=4)
print('   ✅ 已更新 authaddin.json')
"

# 4. 重启 WPS
echo "▶ [4/4] 重启 WPS Office 以加载新插件..."
osascript -e 'tell application "wpsoffice" to quit' 2>/dev/null || true
sleep 3
open -a "/Applications/wpsoffice.app"

echo ""
echo "==================================="
echo " ✅ PPT 插件安装完成！"
echo ""
echo " 请打开或新建一个 PPT 文件,"
echo " 在顶部工具栏找到 [Claude AI] 按钮。"
echo ""
echo " ※ Excel 插件 (3001/5173) 不受影响"
echo "==================================="
echo ""
echo " 前端地址: http://localhost:5174"
echo " 代理地址: http://localhost:3002"
echo ""
echo " 按 Ctrl+C 停止 PPT 插件服务"

trap "kill $PROXY_PID $VITE_PID 2>/dev/null; exit" INT TERM
wait
