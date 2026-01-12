#!/bin/bash
# Script tự động download build mới nhất và copy vào folder build/

echo "🔍 Kiểm tra build mới nhất..."
cd /workspaces/Quoc_MEP

# Lấy ID của build mới nhất
RUN_ID=$(gh run list --limit 1 --json databaseId --jq '.[0].databaseId')
echo "📦 Build ID: $RUN_ID"

# Xóa folder tạm
rm -rf build_output

# Download build
echo "⬇️  Đang download build..."
mkdir -p build_output
gh run download $RUN_ID --dir build_output

# Tạo folder build nếu chưa có
mkdir -p build

# Copy file zip vào folder build
echo "📁 Copy file vào folder build/..."
cp build_output/Quoc_MEP_Universal_Package/*.zip build/

# Hiển thị kết quả
echo "✅ Xong! File đã được copy vào folder build/:"
ls -lh build/*.zip

# Cleanup
echo "🧹 Dọn dẹp folder tạm..."
rm -rf build_output

echo "✨ Hoàn tất! File zip trong: build/"
