#!/usr/bin/env bash
set -euo pipefail

usage() {
  echo "Usage: ./install-macos.sh [--scope user|system] [--skip-debug]"
}

scope="user"
skip_debug="0"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --scope)
      if [[ $# -lt 2 ]]; then
        usage
        exit 1
      fi
      scope="$2"
      shift 2
      ;;
    --skip-debug)
      skip_debug="1"
      shift
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      usage
      exit 1
      ;;
  esac
done

if [[ "$scope" != "user" && "$scope" != "system" ]]; then
  echo "Invalid scope: $scope"
  usage
  exit 1
fi

if [[ "$scope" == "system" && "${EUID}" -ne 0 ]]; then
  echo "System scope requires sudo/root."
  exit 1
fi

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo_root="${script_dir}"
extension_name="PremiereGridMaker"

if [[ "$scope" == "system" ]]; then
  base_path="/Library/Application Support/Adobe/CEP/extensions"
else
  base_path="${HOME}/Library/Application Support/Adobe/CEP/extensions"
fi

install_path="${base_path}/${extension_name}"

mkdir -p "$base_path"
rm -rf "$install_path"
mkdir -p "$install_path"

rsync -a \
  --delete \
  --exclude ".git" \
  --exclude "node_modules" \
  "${repo_root}/" \
  "${install_path}/"

if [[ "$skip_debug" == "0" ]]; then
  for version in 8 9 10 11; do
    defaults write "com.adobe.CSXS.${version}" PlayerDebugMode 1
  done
fi

echo "Installed '${extension_name}' to: ${install_path}"
if [[ "$skip_debug" == "1" ]]; then
  echo "Skipped CEP debug mode changes."
else
  echo "CEP debug mode enabled for com.adobe.CSXS.8 to com.adobe.CSXS.11."
fi
echo "Open Premiere Pro: Window > Extensions (Legacy) > Premiere Grid Maker"
