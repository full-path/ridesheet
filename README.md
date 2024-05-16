# RideSheet Docs

*Built with Mkdocs*

## Installation

These docs are built with Mkdocs, which is a python static site generator. To get started, you will need python3 and pip installed on your machine. To get the necessary dependencies:

```bash
pip install mkdocs pymdown-extensions mkdocs-material
```

From the main folder, you should now be able to run a local version of the site:

```bash
mkdocs serve
```

## Deploying

All of the source files are in the `src/` directory. To modify the main menu or other site details, edit the `mkdocs.yml` file. When you build the site, it will rebuild all files into the `docs/` directory.

```bash
mkdocs build
```

Make sure to commit everything in the `docs/` directory as well as the edited source files, and the site should republish as soon as the commit is pushed to GitHub.

## Common Issues

In order to follow GitHub conventions, the default mkdocs folders have been renamed. When referring to any official mkdocs documentation, they may refer to the `docs/` folder as the source directory. Instead, you will use the `src/` directory for any source material.

