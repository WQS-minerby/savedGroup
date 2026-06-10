param(
  [string]$Remote = "https://github.com/<your-username>/<repo-name>.git",
  [string]$Branch = "main"
)

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  Write-Error "Git is not installed. Install Git from https://git-scm.com/download/win and rerun this script."
  exit 1
}

Set-Location $PSScriptRoot

if (-not (Test-Path .git)) {
  git init
}

git add .
git commit -m "Initial commit" --allow-empty

git branch -M $Branch

git remote remove origin 2>$null

git remote add origin $Remote

git push -u origin $Branch
