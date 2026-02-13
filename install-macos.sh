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

required_items=(
  "CSXS"
  "css"
  "js"
  "jsx"
  "index.html"
)

for item in "${required_items[@]}"; do
  src="${repo_root}/${item}"
  if [[ ! -e "$src" ]]; then
    echo "Missing required item: ${item}"
    exit 1
  fi
done

for item in "${required_items[@]}"; do
  rsync -a --exclude ".DS_Store" "${repo_root}/${item}" "${install_path}/"
done

if [[ "$skip_debug" == "0" ]]; then
  debug_warn="0"
  for version in $(seq 8 15); do
    domain="com.adobe.CSXS.${version}"
    if ! defaults write "${domain}" PlayerDebugMode 1 >/dev/null 2>&1; then
      echo "Warning: failed to set PlayerDebugMode for ${domain}"
      debug_warn="1"
      continue
    fi
    readback="$(defaults read "${domain}" PlayerDebugMode 2>/dev/null || true)"
    if [[ "${readback}" != "1" ]]; then
      echo "Warning: PlayerDebugMode readback is not 1 for ${domain} (got: ${readback:-<empty>})"
      debug_warn="1"
    fi
  done
  if [[ "${debug_warn}" == "1" ]]; then
    echo "Warning: some CEP debug keys could not be validated. Panel may stay blank until PlayerDebugMode=1 is applied."
  fi
fi

echo "Installed runtime files for '${extension_name}' to: ${install_path}"
if [[ "$skip_debug" == "1" ]]; then
  echo "Skipped CEP debug mode changes."
else
  echo "CEP debug mode enabled for com.adobe.CSXS.8 to com.adobe.CSXS.15."
fi
echo "Open Premiere Pro: Window > Extensions > Grid Maker"
