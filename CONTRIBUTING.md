# Contributing Guidelines

Thank you for your interest in contributing to my projects! This document provides general guidelines for contributing to any of my repositories.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [How to Contribute](#how-to-contribute)
- [Coding Standards](#coding-standards)
- [Commit Guidelines](#commit-guidelines)
- [Pull Request Process](#pull-request-process)
- [Issue Reporting](#issue-reporting)
- [Questions and Support](#questions-and-support)

## Code of Conduct

By participating in this project, you agree to abide by the Code of Conduct. Please be respectful, constructive, and professional in all interactions.

## Getting Started

### Prerequisites

Before contributing, ensure you have:

1. **Git** installed and configured
2. **GitHub account** set up
3. **Development environment** appropriate for the project (Python, Node.js, etc.)
4. **Forked the repository** to your own GitHub account

### First Time Setup

1. Fork the repository
2. Clone your fork locally:
   ```bash
   git clone https://github.com/YOUR-USERNAME/REPOSITORY-NAME.git
   cd REPOSITORY-NAME
   ```
3. Add the upstream repository:
   ```bash
   git remote add upstream https://github.com/ORIGINAL-OWNER/REPOSITORY-NAME.git
   ```
4. Create a branch for your changes:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## How to Contribute

### Types of Contributions

I welcome various types of contributions:

- 🐛 **Bug fixes** - Fix issues and improve stability
- ✨ **New features** - Add functionality that aligns with project goals
- 📝 **Documentation** - Improve README, comments, or guides
- 🎨 **UI/UX improvements** - Enhance user interface and experience
- ⚡ **Performance** - Optimize code and reduce resource usage
- 🧪 **Tests** - Add or improve test coverage
- 🌍 **Translations** - Add language support (if applicable)

### Before You Start

1. **Check existing issues** - Someone might already be working on it
2. **Open an issue first** - For significant changes, discuss your approach
3. **Keep changes focused** - One feature/fix per pull request
4. **Update documentation** - If your changes affect usage

## Coding Standards

### General Principles

- ✅ **Write clean, readable code** - Code is read more often than written
- ✅ **Follow existing style** - Match the style of the project
- ✅ **Keep it simple** - Avoid unnecessary complexity
- ✅ **Comment when needed** - Explain *why*, not *what*
- ✅ **Test your changes** - Ensure nothing breaks

### Python Projects

For Python-based projects, please follow these guidelines:

#### Code Style

- **PEP 8 compliance** - Follow Python's style guide
- **Use `#` comments** - Not `"""` docstrings for inline comments
- **UTF-8 encoding** - Include `# -*- coding: utf-8 -*-` at the top of files
- **Preserve Unicode characters** - Maintain UI symbols (★, ◈, ▶, etc.) in UTF-8

#### Best Practices

```python
# Good - Clear variable names, proper spacing
def calculate_mining_yield(ore_volume, cycle_time):
    # Calculate theoretical yield per hour
    cycles_per_hour = 3600 / cycle_time
    return ore_volume * cycles_per_hour

# Bad - Unclear names, no comments
def calc(v,t):
    return v*(3600/t)
```

#### What NOT to Change

Unless explicitly required for the task:

- ❌ **Core logic/algorithms** - Don't modify calculation logic without discussion
- ❌ **File formats** - Maintain existing comment styles and structures
- ❌ **Visual layouts** - Keep UI arrangements consistent
- ❌ **Working features** - If it ain't broke, don't fix it

### JavaScript/TypeScript Projects

- Use **consistent indentation** (2 or 4 spaces as per project)
- Follow **ESLint rules** if configured
- Use **meaningful variable names**
- Add **JSDoc comments** for functions

### Other Languages

- Follow the **established patterns** in the project
- Use **language-specific linters** when available
- Maintain **consistent formatting**

## Commit Guidelines

### Commit Message Format

Use clear, descriptive commit messages:

```
<type>: <short description>

<optional longer description>

<optional footer>
```

### Types

- `feat:` - New feature
- `fix:` - Bug fix
- `docs:` - Documentation changes
- `style:` - Code style changes (formatting, no logic change)
- `refactor:` - Code refactoring
- `perf:` - Performance improvements
- `test:` - Adding or updating tests
- `chore:` - Maintenance tasks

### Examples

```bash
# Good
git commit -m "fix: Correct residue calculation for compressed ore"
git commit -m "feat: Add Discord webhook integration for fleet reports"
git commit -m "docs: Update installation instructions for Windows users"

# Bad
git commit -m "fixed stuff"
git commit -m "changes"
git commit -m "asdf"
```

### Keep Commits Atomic

- One logical change per commit
- Commits should be self-contained
- Easy to review and revert if needed

## Pull Request Process

### Before Submitting

1. ✅ **Update from upstream** - Sync with the latest changes
   ```bash
   git fetch upstream
   git rebase upstream/main
   ```

2. ✅ **Test thoroughly** - Ensure everything works
   - Run the application
   - Test your specific changes
   - Check for regressions

3. ✅ **Update documentation** - If applicable
   - README.md updates
   - Code comments
   - Usage examples

4. ✅ **Clean up commits** - Squash if needed
   ```bash
   git rebase -i HEAD~3  # Interactive rebase last 3 commits
   ```

### Submitting the PR

1. Push to your fork:
   ```bash
   git push origin feature/your-feature-name
   ```

2. Open a Pull Request on GitHub

3. Fill out the PR template with:
   - **Clear title** - Summarize the change
   - **Description** - What and why
   - **Testing done** - How you verified it works
   - **Screenshots** - If UI changes
   - **Related issues** - Link to issue numbers

### PR Template Example

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Documentation update
- [ ] Performance improvement

## Testing
- [ ] Tested locally
- [ ] No regressions found
- [ ] Edge cases considered

## Screenshots (if applicable)
[Add screenshots here]

## Related Issues
Fixes #123
```

### Review Process

- I'll review PRs as time permits (usually within a week)
- Be patient and respectful
- Address feedback constructively
- Make requested changes in new commits, then squash when approved

### After Approval

- I'll merge your PR
- Your contribution will be credited
- Delete your feature branch (optional but clean)

## Issue Reporting

### Before Opening an Issue

1. **Search existing issues** - Avoid duplicates
2. **Use latest version** - Bug might be fixed already
3. **Gather information** - Logs, screenshots, steps to reproduce

### Bug Report Template

```markdown
## Bug Description
Clear description of the bug

## Steps to Reproduce
1. Go to '...'
2. Click on '...'
3. See error

## Expected Behavior
What should happen

## Actual Behavior
What actually happens

## Environment
- OS: Windows 10
- Python Version: 3.11
- Application Version: 1.2.3

## Logs/Screenshots
[Paste relevant logs or add screenshots]

## Additional Context
Any other relevant information
```

### Feature Request Template

```markdown
## Feature Description
Clear description of the proposed feature

## Use Case
Why is this feature needed? What problem does it solve?

## Proposed Solution
How should this work?

## Alternatives Considered
What other approaches did you think about?

## Additional Context
Mockups, examples, references
```

## Questions and Support

### Where to Ask

- 💬 **GitHub Discussions** - General questions and discussions (if enabled)
- 🐛 **GitHub Issues** - Bug reports and feature requests
- 📧 **Email** - For private/security concerns (see repository for contact)

### Getting Help

- Be specific about your problem
- Include relevant details (OS, version, logs)
- Be patient and respectful
- Help others when you can

## Project-Specific Guidelines

Some repositories may have additional guidelines in their own `CONTRIBUTING.md`. Always check the project's documentation for:

- **Technology-specific requirements** (Python version, dependencies)
- **Build/test instructions**
- **Special coding conventions**
- **Domain-specific rules** (e.g., EVE Online EULA compliance)

## Recognition

Contributors will be acknowledged in:

- Release notes
- CONTRIBUTORS.md file (if applicable)
- Project documentation

Significant contributors may be offered maintainer status for ongoing projects.

## License

By contributing, you agree that your contributions will be licensed under the same license as the project (typically specified in LICENSE file).

---

## Quick Checklist

Before submitting a PR, verify:

- [ ] Code follows project style guidelines
- [ ] UTF-8 encoding preserved (no garbage characters)
- [ ] Comments use `#` format (for Python projects)
- [ ] Core logic unchanged unless required
- [ ] All tests pass (if applicable)
- [ ] Documentation updated
- [ ] Commits are clear and atomic
- [ ] PR description is complete

---

Thank you for contributing! Your time and effort help make these projects better for everyone.

**Questions?** Open an issue and tag it with `question` label.

*Happy coding!* 🚀
