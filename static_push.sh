#!/bin/bash

# 构建Jekyll站点
bundle exec jekyll build

# 进入子模块目录
cd _site

# 检查是否有更改
if [[ -n $(git status -s) ]]; then
  # 先暂存所有更改
  git add .
  
  # 拉取远程更新，如有冲突会暂停让用户解决
  if ! git pull origin main --rebase; then
    echo "Rebase encountered conflicts. Please resolve them and run 'git rebase --continue' in _site directory, then re-run this script."
    exit 1
  fi
  
  # 提交更改
  git commit -m "Update site: $(date +'%Y-%m-%d %H:%M:%S')"
  
  # 推送更改
  if ! git push origin main; then
    echo "Failed to push to _site repository. Please check manually."
    exit 1
  fi
else
  echo "No changes to commit in _site"
fi

# 返回主目录
cd ..

# 检查主仓库当前分支
current_branch=$(git rev-parse --abbrev-ref HEAD)
echo "Current branch in main repo: $current_branch"

# 检查主仓库是否有子模块引用更新
if [[ -n $(git status -s _site) ]]; then
  # 提交子模块引用更新
  git add _site
  git commit -m "Update submodule reference: $(date +'%Y-%m-%d %H:%M:%S')"
  
  # 推送更改，使用当前分支
  if ! git push origin $current_branch; then
    echo "Failed to push to main repository. Please check manually."
    exit 1
  fi
else
  echo "No submodule reference changes to commit"
fi
