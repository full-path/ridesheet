# Updating the Code

This page covers setting up a local development environment using GitHub and Clasp, how to update the code when new RideSheet versions are released, and how to contribute back to the main codebase.

## Required Tools

Before getting started, you'll need to install and configure these tools:

- [Git](https://git-scm.com/) - Version control system for tracking code changes
- [GitHub Account](https://github.com/) - For hosting your code and collaborating with others
- [Clasp](https://github.com/google/clasp) - Google's Command Line Apps Script Projects tool for developing Apps Script locally

## Setting Up Your Development Environment

Follow these steps to set up a local development environment:

1. Fork the [RideSheet repository](https://github.com/full-path/ridesheet) on GitHub to create your own copy
2. Get your Apps Script project ID:
    - Open your RideSheet spreadsheet
    - Go to `Extensions > Apps Script`
    - Click Project Settings
    - Copy the Script ID
3. Clone your Apps Script project locally (this will create a new folder):
    ```bash
    clasp clone <your-script-id>
    ```
4. Initialize Git in the folder created by clasp:
    ```bash
    cd <project-folder>
    git init
    ```
5. Connect to your forked repository:
    ```bash
    git remote add origin <your-fork-url>
    git fetch origin
    git reset --hard origin/main
    ```
6. Push the main branch code to your RideSheet instance:
    ```bash
    clasp push
    ```

## Making Changes

When making your own customizations on your forked repository:

1. Create a new branch for your changes:
    ```bash
    git checkout -b my-custom-feature
    ```
2. Make changes to the code locally
3. Push changes to Apps Script to test them:
    ```bash
    clasp push
    ```
4. Test your changes in the Apps Script editor and RideSheet spreadsheet
5. Once satisfied, commit and push to GitHub:
    ```bash
    git add .
    git commit -m "Description of changes"
    git push origin my-custom-feature
    ```

!!! note "Note"
    If you're only planning to receive updates from the main RideSheet repository, you don't need to create feature branches.

## Contributing Back

RideSheet welcomes contributions from the community! We accept pull requests for:

- Bug fixes
- New features with broad applicability
- Improvements to existing functionality

Before starting work on a major feature, please:

1. [Open an issue](https://github.com/full-path/ridesheet/issues/new) to discuss the proposed changes
2. Contact the RideSheet team for guidance
3. Ensure your contribution would benefit multiple agencies

This helps ensure your time is well spent and the feature aligns with RideSheet's goals.

