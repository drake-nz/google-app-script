# GitHub Setup Instructions

Follow these steps to upload this project to GitHub as a separate repository.

## Step 1: Create a New GitHub Repository

1. Go to [GitHub](https://github.com) and sign in
2. Click the **"+"** icon in the top right → **"New repository"**
3. Repository name: `google-apps-script-projects` (or your preferred name)
4. Description: "Google Apps Script projects for Sheets automation with rich text formatting"
5. Choose **Public** or **Private**
6. **DO NOT** initialize with README, .gitignore, or license (we already have these)
7. Click **"Create repository"**

## Step 2: Initialize Git in Your Project

Open terminal and navigate to your project:

```bash
cd /Users/dot/Scripts/apps-script-projects
```

## Step 3: Initialize Git Repository

```bash
# Initialize git (if not already done)
git init

# Check current status
git status
```

## Step 4: Add All Files

```bash
# Add all files (respects .gitignore)
git add .

# Verify what will be committed
git status
```

**Important:** Make sure `.clasp.json` files are NOT being added (they should be ignored by `.gitignore`)

## Step 5: Create Initial Commit

```bash
git commit -m "Initial commit: Google Apps Script projects for rich text concatenation"
```

## Step 6: Connect to GitHub

Replace `YOUR_USERNAME` and `YOUR_REPO_NAME` with your actual GitHub username and repository name:

```bash
# Add remote repository
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git

# Or if using SSH:
# git remote add origin git@github.com:YOUR_USERNAME/YOUR_REPO_NAME.git

# Verify remote
git remote -v
```

## Step 7: Push to GitHub

```bash
# Push to main branch
git branch -M main
git push -u origin main
```

## Step 8: Verify

1. Go to your GitHub repository page
2. Verify all files are present
3. Check that `.clasp.json` files are NOT visible (they should be ignored)

## Future Updates

When you make changes:

```bash
cd /Users/dot/Scripts/apps-script-projects

# Check what changed
git status

# Add changes
git add .

# Commit
git commit -m "Description of your changes"

# Push
git push
```

## Security Checklist

Before pushing, ensure:

- ✅ `.clasp.json` files are in `.gitignore` (contains personal auth)
- ✅ No API keys or secrets in code
- ✅ Sheet IDs are in `Config.gs` (consider using environment variables for sensitive IDs)
- ✅ `.clasprc.json` is ignored (if it exists)

## Troubleshooting

**"Repository not found":**
- Check your GitHub username and repository name
- Verify you have push access

**"Authentication failed":**
- Use GitHub Personal Access Token instead of password
- Or set up SSH keys

**"Files not showing up":**
- Check `.gitignore` isn't excluding too much
- Run `git status` to see what's tracked

**"Want to remove .clasp.json from tracking":**
```bash
git rm --cached .clasp.json
git commit -m "Remove .clasp.json from tracking"
```
