# Contributing to Microsoft 365 MCP Server

Thank you for your interest in contributing to the Microsoft 365 MCP Server project.

## Development Setup

### Prerequisites
- Node.js 18+
- Microsoft 365 Business/Enterprise account with admin access
- Cloudflare Workers account
- Git

### Local Development
1. Fork and clone the repository
2. Install dependencies: `npm install`
3. Copy configuration files:
   ```bash
   cp .dev.vars.example .dev.vars
   cp wrangler.example.toml wrangler.toml
   ```
4. Configure your Microsoft 365 and Cloudflare credentials
5. Start development server: `npm run dev`

## Code Standards

### TypeScript
- Follow existing TypeScript configurations
- Ensure type safety with `npm run type-check`
- Use proper error handling patterns

### Code Quality
- Run linting: `npm run lint`
- Format code: `npm run format`
- Follow existing code style and patterns

### Documentation
- Update relevant documentation for new features
- Include code examples where appropriate
- Maintain professional tone in all documentation

## Testing

Currently, the project has minimal testing infrastructure. When contributing:
- Manually test all changes thoroughly
- Verify OAuth flows work correctly
- Test Microsoft Graph API integrations
- Ensure no breaking changes to existing functionality

## Submission Process

### Pull Requests
1. Create a feature branch from `main`
2. Make your changes with clear, focused commits
3. Update documentation as needed
4. Test your changes thoroughly
5. Submit a pull request with:
   - Clear description of changes
   - Motivation for the changes
   - Testing steps performed

### Commit Messages
- Use clear, descriptive commit messages
- Focus on what the change accomplishes
- Keep messages concise but informative

## Code of Conduct

### Professional Standards
- Maintain professional communication
- Respect other contributors' work and opinions
- Focus on constructive feedback and solutions
- Follow project coding standards consistently

### Security Considerations
- Never commit secrets or credentials
- Use placeholder values in examples
- Follow security best practices
- Report security issues privately

## Areas for Contribution

### High Priority
- Unit and integration test implementation
- Additional Microsoft Graph API tools
- Error handling improvements
- Performance optimizations

### Medium Priority  
- Documentation improvements
- Code quality enhancements
- Development workflow improvements
- Example applications

### Low Priority
- Advanced features and extensions
- Alternative authentication methods
- Additional platform integrations

## Getting Help

- Review existing documentation in this repository
- Check existing issues and discussions
- Create an issue for questions or bug reports
- Provide detailed information when reporting issues

## License

By contributing to this project, you agree that your contributions will be licensed under the same MIT License that covers the project.