#!/usr/bin/env bash
# This file can be double-clicked on macOS to run the push script

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Run the push script
"$SCRIPT_DIR/push_to_github.sh"

# Keep terminal open so user can see the result
echo ""
echo "Press any key to close this window..."
read -n 1 -s

