#!/usr/bin/env bash
set -euo pipefail

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

# Change to project directory
cd "$PROJECT_DIR"

echo "🚀 Pushing updates to GitHub..."
echo "📁 Project directory: $PROJECT_DIR"
echo ""

# Check if we're in a git repository
if ! git rev-parse --git-dir > /dev/null 2>&1; then
    echo "❌ Error: Not a git repository!"
    exit 1
fi

# Check if there are any changes
if git diff --quiet && git diff --cached --quiet; then
    echo "ℹ️  No changes to commit. Working tree is clean."
    exit 0
fi

# Show status
echo "📊 Current status:"
git status --short
echo ""

# Prompt for commit message
if [ -t 0 ]; then
    # Interactive mode - prompt for commit message
    echo "💬 Enter commit message (or press Enter for default):"
    read -r COMMIT_MSG
    if [ -z "$COMMIT_MSG" ]; then
        COMMIT_MSG="Update project files"
    fi
else
    # Non-interactive mode - use default
    COMMIT_MSG="Update project files"
fi

# Add all changes
echo "➕ Staging all changes..."
git add .

# Commit changes
echo "💾 Committing changes..."
git commit -m "$COMMIT_MSG"

# Push to GitHub
echo "📤 Pushing to GitHub..."
git push

echo ""
echo "✅ Successfully pushed to GitHub!"
echo "🔗 Repository: $(git remote get-url origin)"

