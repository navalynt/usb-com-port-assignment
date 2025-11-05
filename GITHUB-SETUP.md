# GitHub Repository Setup Instructions

This guide will help you upload the USB COM Port Assignment project to your GitHub repository.

## Quick Start

### Option 1: Using GitHub Web Interface (Easiest)

1. **Create a new repository on GitHub:**
   - Go to https://github.com/new
   - Repository name: `usb-com-port-assignment` (or your choice)
   - Description: "Automated COM port assignment for USB devices in Omnissa Horizon VDI"
   - Choose Public or Private
   - **DO NOT** initialize with README, .gitignore, or license (we have these)
   - Click "Create repository"

2. **Upload files:**
   - Download the entire `github-export` folder from Claude
   - On your new repository page, click "uploading an existing file"
   - Drag and drop ALL files and folders from `github-export`
   - Commit message: "Initial commit - Version 1.8"
   - Click "Commit changes"

### Option 2: Using Git Command Line

1. **Create repository on GitHub** (follow step 1 above)

2. **Download and extract** the `github-export` folder

3. **Initialize and push:**
   ```bash
   cd github-export
   
   git init
   git add .
   git commit -m "Initial commit - Version 1.8"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
   git push -u origin main
   ```

### Option 3: Using GitHub Desktop

1. **Create repository on GitHub** (follow step 1 above)

2. **Download** the `github-export` folder

3. **In GitHub Desktop:**
   - File → Clone Repository
   - Select your new repository
   - Choose location
   - Copy all files from `github-export` into the cloned folder
   - Commit to main: "Initial commit - Version 1.8"
   - Push to origin

## Repository Structure

Your repository will have this structure:

```
usb-com-port-assignment/
├── .github/
│   └── workflows/
│       └── powershell-lint.yml     # GitHub Actions workflow
├── docs/
│   ├── COM-Settings-Quick-Reference.md
│   ├── USB-COM-Deployment-Guide.md
│   ├── USB-COM-Port-Assignment-Documentation.md
│   └── VERSION-1.8-UPDATE.md
├── scripts/
│   ├── Debug-USBCOMDevices.ps1
│   └── Set-USBCOMPortAssignment.ps1
├── .gitignore
├── CHANGELOG.md
├── CONTRIBUTING.md
├── LICENSE
└── README.md
```

## Post-Upload Configuration

### 1. Update README.md

Replace placeholder URLs in README.md:
- Change `YOUR-USERNAME` to your GitHub username
- Change `YOUR-REPO` to your repository name

Example:
```markdown
<!-- Before -->
https://github.com/YOUR-USERNAME/YOUR-REPO/issues

<!-- After -->
https://github.com/john-doe/usb-com-port-assignment/issues
```

### 2. Configure Repository Settings

**In your GitHub repository settings:**

1. **General:**
   - ✅ Enable Issues
   - ✅ Enable Discussions (optional but recommended)
   - ✅ Enable Projects (optional)

2. **Topics/Tags:** Add relevant topics:
   - `powershell`
   - `vdi`
   - `omnissa`
   - `horizon-view`
   - `com-ports`
   - `usb-redirection`
   - `windows`

3. **Description:** 
   ```
   Automated COM port assignment for USB devices in Omnissa Horizon VDI environments. PowerShell 5.0+, UTF-8 compliant, non-persistent desktop optimized.
   ```

4. **Website:** (optional)
   - Link to your organization or documentation site

### 3. Create First Release

1. Go to "Releases" in your repository
2. Click "Create a new release"
3. Tag version: `v1.8.0`
4. Release title: `Version 1.8.0 - Critical Registry Path Fix`
5. Description:
   ```markdown
   ## Critical Fix Release
   
   This release fixes a critical bug where COM settings were not being applied to devices.
   
   ### What's Fixed
   - Registry path resolution in `Set-COMPortSettings` and `Set-COMPortAssignment`
   - COM settings now correctly apply (baud rate, parity, buffers)
   - Devices use manufacturer specifications instead of Windows defaults
   
   ### Supported Devices
   - Topaz T-LBK462-BSB-RC → COM5 (19200 baud, Odd parity)
   - Ingenico LANE3000 → COM21 (115200 baud)
   
   ### Installation
   See [README.md](README.md) for complete installation instructions.
   
   ### Upgrading from 1.7.0 or earlier
   This is a critical fix. Previous versions assigned COM ports but did not apply custom settings.
   Replace the script on your gold image and redeploy.
   ```
6. Attach files (optional): You can attach zip files of the scripts
7. Click "Publish release"

### 4. Enable GitHub Actions

The repository includes a GitHub Actions workflow for automatic linting.

1. Go to "Actions" tab in your repository
2. You'll see "PowerShell Linting and Validation"
3. It will run automatically on push/PR
4. (Optional) Add a badge to README.md:
   ```markdown
   [![CI](https://github.com/YOUR-USERNAME/YOUR-REPO/actions/workflows/powershell-lint.yml/badge.svg)](https://github.com/YOUR-USERNAME/YOUR-REPO/actions)
   ```

### 5. Set Up GitHub Pages (Optional)

For documentation hosting:

1. Settings → Pages
2. Source: Deploy from a branch
3. Branch: main
4. Folder: /docs
5. Save

Your documentation will be available at:
`https://YOUR-USERNAME.github.io/YOUR-REPO/`

## Recommended GitHub Repository Settings

### Branch Protection Rules (for main branch)

Settings → Branches → Add rule:
- Branch name pattern: `main`
- ✅ Require pull request reviews before merging
- ✅ Require status checks to pass before merging
  - ✅ PowerShell Linting and Validation
- ✅ Require branches to be up to date before merging

### Labels

Add these labels for issue tracking:
- `bug` - Something isn't working
- `enhancement` - New feature or request
- `documentation` - Documentation improvements
- `good first issue` - Good for newcomers
- `help wanted` - Extra attention needed
- `question` - Further information requested
- `device-support` - Adding support for new devices

### Issue Templates

Create issue templates in `.github/ISSUE_TEMPLATE/`:

**Bug Report Template:**
```markdown
---
name: Bug Report
about: Report a problem with the script
---

## Description
A clear description of the bug.

## Environment
- Script Version: [e.g., 1.8.0]
- Windows Version: [e.g., Windows 11 22H2]
- PowerShell Version: [e.g., 5.1]
- VDI Platform: [e.g., Omnissa Horizon View 2506]

## Device Information
- Device Model:
- VendorID:
- ProductID:
- Target COM Port:

## Steps to Reproduce
1. 
2. 
3. 

## Expected Behavior


## Actual Behavior


## Log Output
```
Paste relevant log entries here
```

## Additional Context
```

**Feature Request Template:**
```markdown
---
name: Feature Request
about: Suggest a new feature or enhancement
---

## Problem Description
What problem does this feature solve?

## Proposed Solution
How should this work?

## Alternatives Considered
What other solutions did you consider?

## Additional Context
```

## Sharing Your Repository

Once your repository is set up:

1. **Share the URL:**
   ```
   https://github.com/YOUR-USERNAME/YOUR-REPO
   ```

2. **Installation command becomes:**
   ```powershell
   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/YOUR-USERNAME/YOUR-REPO/main/scripts/Set-USBCOMPortAssignment.ps1" `
       -OutFile "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1"
   ```

3. **Documentation links:**
   ```
   https://github.com/YOUR-USERNAME/YOUR-REPO/blob/main/docs/
   ```

## Maintenance Tips

### Regular Updates

1. **Update CHANGELOG.md** for each release
2. **Tag versions** following semantic versioning
3. **Update documentation** when adding features
4. **Respond to issues** promptly
5. **Review and merge PRs** from contributors

### Security

1. **Enable Dependabot** (Settings → Security → Dependabot)
2. **Review security advisories** regularly
3. **Keep dependencies updated** (if any)
4. **Never commit sensitive data** (credentials, keys)

### Community

1. **Welcome contributors** in CONTRIBUTING.md
2. **Use GitHub Discussions** for Q&A
3. **Create a Wiki** for extended documentation
4. **Add a CODE_OF_CONDUCT.md** for larger projects

## Troubleshooting

### Upload Issues

**"Push rejected":**
- Make sure you don't have README/LICENSE already in repo
- Force push: `git push -f origin main` (only on initial commit)

**"Large files":**
- GitHub has 100MB file limit
- All our files are well under this

**"Permission denied":**
- Check your authentication (token or SSH key)
- Verify repository permissions

### GitHub Actions Not Running

- Check "Actions" tab is enabled in repository settings
- Workflow file must be in `.github/workflows/`
- Workflow must have `.yml` or `.yaml` extension

## Next Steps

After setting up your repository:

1. ✅ Share with your team
2. ✅ Create first release (v1.8.0)
3. ✅ Update README with your repo URLs
4. ✅ Test the GitHub Actions workflow
5. ✅ Add repository topics/tags
6. ✅ (Optional) Enable GitHub Discussions
7. ✅ (Optional) Set up GitHub Pages for docs

## Support

If you need help:
- GitHub Docs: https://docs.github.com
- GitHub Community: https://github.community
- Git Handbook: https://guides.github.com/introduction/git-handbook/

---

**You're all set! Your professional PowerShell project is ready for GitHub!** 🚀
