# Copilot Instructions

## Documentation Requirements
All functions should be documented in English. All parameters should be documented as well.

- Use comprehensive docstrings for all functions and classes
- Document parameter types and return values
- Include usage examples when helpful
- Follow Google or NumPy docstring style

## Coding Style Requirements
Follow the settings in flake8. Reformat files as necessary. Use type hints. Stop and ask when there is ambiguity.

- Apply type hints to all function parameters and return values
- Follow PEP 8 style guidelines
- Use flake8 configuration for linting rules
- Format code with Black or similar formatter
- Ask for clarification when requirements are unclear or ambiguous
- Strings should be enclosed using double quotes unless single quotes are necessary to avoid escaping.

## Testing Requirements
Write code that is easy to test and include test cases when appropriate.

- Structure code to enable unit testing (avoid tightly coupled functions)
- Separate business logic from I/O operations when possible
- Include test cases for new functions and critical code paths
- Use pytest conventions for test structure and naming
- Include both positive and negative test cases
- Mock external dependencies (files, APIs, databases) in tests
- Write testable code by using dependency injection where appropriate
