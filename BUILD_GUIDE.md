# Build Management

## Folder Structure

```
build/                          # Folder chứa tất cả builds
├── Quoc_MEP_20260112_093219.zip
├── Quoc_MEP_20260112_101234.zip
└── ...
```

## Quick Commands

### Download build mới nhất
```bash
./download-latest-build.sh
```

Script này sẽ:
1. Tìm build mới nhất trên GitHub Actions
2. Download về
3. Copy file zip vào folder `build/`
4. Tự động dọn dẹp folder tạm

### Kiểm tra builds có sẵn
```bash
ls -lh build/
```

### Xóa builds cũ (giữ lại 5 builds gần nhất)
```bash
cd build && ls -t | tail -n +6 | xargs rm -f
```

## Manual Download

Nếu muốn download thủ công:
```bash
# Lấy ID của build
gh run list --limit 1

# Download build cụ thể (thay RUN_ID)
gh run download RUN_ID -n Quoc_MEP_Universal_Package -D ./temp_download
cp ./temp_download/*.zip build/
rm -rf ./temp_download
```
