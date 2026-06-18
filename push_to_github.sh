#!/bin/bash
# 一键推送到 GitHub — 网络恢复后运行此脚本
cd "$(dirname "$0")"
git remote set-url origin "https://xubangkun439-create:$(cat .gh_token 2>/dev/null || echo 'YOUR_TOKEN')@github.com/xubangkun439-create/CorelDRAWer-Skill.git"
git push origin main
git remote set-url origin https://github.com/xubangkun439-create/CorelDRAWer-Skill.git
echo "✅ Push done: https://github.com/xubangkun439-create/CorelDRAWer-Skill"
