#!/bin/bash
# Setup script for all Apps Script projects
# This script initializes all projects in the workspace

set -e

echo "🚀 Setting up Multi-Project Apps Script Workspace"
echo "=================================================="
echo ""

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "❌ clasp is not installed."
    echo "   Install it with: npm install -g @google/clasp"
    exit 1
fi

echo "✅ clasp is installed: $(clasp --version)"
echo ""

# Check login status
echo "Checking login status..."
if clasp login --status &> /dev/null; then
    echo "✅ You are logged in to Google"
else
    echo "⚠️  You are not logged in"
    echo ""
    read -p "Do you want to login now? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo "Opening browser for authentication..."
        clasp login
    else
        echo "Please run 'clasp login' before continuing"
        exit 1
    fi
fi

echo ""
echo "📦 Setting up projects..."
echo ""

# Function to setup a project
setup_project() {
    local project_dir=$1
    local project_title=$2
    
    if [ ! -d "$project_dir" ]; then
        echo "⚠️  Directory $project_dir not found, skipping..."
        return
    fi
    
    cd "$project_dir"
    
    if [ -f ".clasp.json" ]; then
        echo "✅ $project_title already initialized"
        cd ..
        return
    fi
    
    echo "📝 Setting up $project_title..."
    read -p "Create project '$project_title'? (y/n) " -n 1 -r
    echo ""
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        clasp create --title "$project_title" --type standalone
        if [ -f "Code.gs" ]; then
            clasp push
            echo "✅ $project_title created and code pushed"
        else
            echo "✅ $project_title created (no Code.gs found, skipping push)"
        fi
    else
        echo "⏭️  Skipped $project_title"
    fi
    
    cd ..
}

# Setup each project
BASE_DIR=$(pwd)

setup_project "formatting" "Sheet Formatting"
setup_project "slides-export" "Slides Export"
setup_project "shared-utilities" "Shared Utilities"

echo ""
echo "✨ Setup Complete!"
echo ""
echo "📋 Next Steps:"
echo "=============="
echo ""
echo "1. Add your Sheet IDs to shared-utilities/Config.gs"
echo "2. Test formatting:"
echo "   cd formatting && clasp open-script"
echo ""
echo "3. Test slides export:"
echo "   cd slides-export && clasp open-script"
echo ""
echo "4. (Optional) Set up shared-utilities as a library:"
echo "   - Deploy shared-utilities project"
echo "   - Get Script ID: cd shared-utilities && clasp status"
echo "   - Add as library in other projects via Apps Script editor"
echo ""
echo "📚 See README.md for detailed documentation"
