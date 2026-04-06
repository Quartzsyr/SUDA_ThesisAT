# Auto Git Backup

This project now includes a local backup script:

`powershell -NoProfile -ExecutionPolicy Bypass -File "E:\paperwrite\scripts\auto_git_backup.ps1"`

What it does:

- runs inside `E:\paperwrite`
- checks whether the directory is a git repository
- skips clean working trees
- stages all changes and creates a timestamped commit

Recommended scheduled task command:

`powershell -NoProfile -ExecutionPolicy Bypass -File "E:\paperwrite\scripts\auto_git_backup.ps1"`

Recommended task name:

`Paperwrite Git Backup`
