# GitHub Pages Deployment Guide

## Repository Configuration Steps

Follow these steps to properly configure GitHub Pages for this Outlook add-in:

### 1. Repository Settings
1. Go to your GitHub repository: `https://github.com/ridvan-soyler/outlook-auto-reply-addin-v2`
2. Click on **Settings** tab
3. Scroll down to **Pages** section in the left sidebar

### 2. Configure GitHub Pages Source
1. In the **Pages** settings:
   - **Source**: Select "GitHub Actions"
   - **Branch**: This will be automatically managed by the workflow

### 3. Environment Setup (Important!)
1. Still in **Settings**, go to **Environments** in the left sidebar
2. Click **New environment**
3. Name it exactly: `github-pages`
4. Click **Configure environment**
5. Under **Environment protection rules**:
   - Check "Required reviewers" if you want manual approval (optional)
   - Add any deployment branches if needed (optional)

### 4. Repository Permissions
1. Go to **Settings** > **Actions** > **General**
2. Under **Workflow permissions**:
   - Select "Read and write permissions"
   - Check "Allow GitHub Actions to create and approve pull requests"

### 5. Push Your Code
```bash
git add .
git commit -m "Initial Outlook add-in with GitHub Pages deployment"
git branch -M main
git remote add origin https://github.com/ridvan-soyler/outlook-auto-reply-addin-v2.git
git push -u origin main
```

### 6. Monitor Deployment
1. Go to **Actions** tab in your repository
2. You should see the "Deploy to GitHub Pages" workflow running
3. Once completed, your site will be available at:
   `https://ridvan-soyler.github.io/outlook-auto-reply-addin-v2/`

## Troubleshooting

### If deployment still fails:
1. Check that the `github-pages` environment exists
2. Verify workflow permissions are set correctly
3. Ensure the repository is public (or you have GitHub Pro/Enterprise for private repo Pages)

### Manual trigger:
- Go to **Actions** tab
- Select "Deploy to GitHub Pages" workflow
- Click "Run workflow" button

## Notes
- The `GH_PAT` secret is not needed for this setup as we're using GitHub's built-in GITHUB_TOKEN
- The workflow will automatically run on every push to the main branch
- The `enablement: true` parameter will automatically configure Pages if not already set up
