# Publish the source repository

The prepared `github/FINDEX` folder is the intended commit root. Do not commit the
entire working project, the `release` folder, installer, portable ZIP, virtual
environments, native tools, generated reports, backups or case files.

## First commit

From the prepared FINDEX folder, review the files, then run the commands below.
This workspace's filesystem does not record ownership, so Git's ownership check
requires an explicit trust path. The `-c` option trusts only this exact folder for
each command; it does not change global Git settings or trust other repositories.

```powershell
$repo = (Get-Location).Path.Replace('\','/')
git -c "safe.directory=$repo" status --short
git -c "safe.directory=$repo" add .
git -c "safe.directory=$repo" diff --cached --stat
git -c "safe.directory=$repo" diff --cached --check
git -c "safe.directory=$repo" commit -m "Publish FINDEX 3.0.1 source"
```

Create an empty GitHub repository named FINDEX. Do not initialize a second README
or license remotely when pushing this prepared local repository. Replace the
example username below with your own repository URL:

```powershell
git -c "safe.directory=$repo" remote add origin https://github.com/YOUR-USERNAME/FINDEX.git
git -c "safe.directory=$repo" push -u origin main
```

The preparation initializes a local main branch only. It does not create a commit,
configure your identity, create a remote, push anything or post on LinkedIn.
If Git has no author identity configured, set your chosen name and GitHub-verified
or GitHub noreply email locally for this repository before committing. Do not
publish a private email unintentionally; use your account's real noreply address.

Before announcing, enable private vulnerability reporting in the repository's
Settings → Advanced Security → Private vulnerability reporting. Check GitHub's
current instructions: https://docs.github.com/en/code-security/how-tos/report-and-fix-vulnerabilities/configure-vulnerability-reporting/configure-for-a-repository
Do not claim it is enabled until the setting is actually enabled.

Review the published README and license in GitHub. Use the provided LinkedIn
draft after replacing its repository-link placeholder with the real URL.
Screenshots/video must use synthetic files and conceal personal paths.

## What is ready to share

Original FINDEX source, included project assets, documentation and license texts.
The native-binary redistribution gap does not get waived by noncommercial intent.
This source publication omits those binaries; users provision dependencies locally.

## What is still on hold

Do not attach the earlier local Setup EXE or portable ZIP. Their exact FFmpeg and
transitive source/build materials are not assembled/verified, and Qt binary
findings are not cleared by a Python package scan. See NATIVE_DISTRIBUTION_REVIEW.md.
No legal/security certification is asserted by this source-only preparation.
