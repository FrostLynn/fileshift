#!/bin/bash
set -e

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_DIR"

if [ -d ".git_repo" ]; then
  echo "Menyinkronkan database git dari .git_repo ke .git..."
  cp -r .git_repo/* .git/ 2>/dev/null || cp -a .git_repo/. .git/
  rm -rf .git_repo
  
  cat << 'WRAP_EOF' > bin/git-wrap.sh
#!/bin/bash
exec git "$@"
WRAP_EOF
  chmod +x bin/git-wrap.sh
  echo "Sukses: Seluruh riwayat Git telah disatukan kembali ke .git!"
else
  echo "Direktori .git_repo tidak ditemukan (repositori sudah menggunakan .git)."
fi
