#!/usr/bin/env bash
set -euo pipefail
umask 077
cd "$(dirname "$(realpath "$0")")"
if [[ "${1:-}" == --login ]]; then
  exec .venv/bin/python shrubbery/worker.py --login
fi
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python shrubbery/worker.py --setup
mkdir -p "$HOME/.config/systemd/user"
cat > "$HOME/.config/systemd/user/telegram-shrubbery.service" <<EOF
[Unit]
Description=Shrubbery Telegram and iCloud School updates
[Service]
ExecStart="$(pwd)/.venv/bin/python" "$(pwd)/shrubbery/worker.py"
Restart=on-failure
RestartSec=15
[Install]
WantedBy=default.target
EOF
systemctl --user daemon-reload
systemctl --user enable telegram-shrubbery.service
systemctl --user restart telegram-shrubbery.service
