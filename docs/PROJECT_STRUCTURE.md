# Project Structure

This document describes the complete folder structure and organization of the CHAINSAW PROPOSITURAS project.

## Directory Overview

```text
chainsaw/
├── 📁 assets/                 # Static resources and media files
├── 📁 config/                 # Configuration files  
├── 📁 docs/                   # Project documentation
├── 📁 examples/               # Example documents and use cases
├── 📁 scripts/                # Installation and utility scripts
├── 📁 src/                    # Source code
├── 📄 CHANGELOG.md           # Version history and changes
├── 📄 CODE_OF_CONDUCT.md     # Community guidelines
├── 📄 LICENSE                # Project license
└── 📄 README.md              # Main project documentation
```

## Detailed Structure

### `/assets/` - Resources

Contains static resources used by the project:

```text
assets/
└── stamp.png                 # Institutional logo/stamp for documents
```

**Purpose:** Store images, icons, templates, and other static resources used during document processing.

### `/config/` - Configuration

Configuration files and settings:

```text
config/
├── chainsaw-config.ini       # Main configuration file
└── word/                     # Word-specific configurations
    └── Word Personalizações.exportedUI
```

**Key Files:**
- `chainsaw-config.ini`: Master configuration with 100+ settings across 15 categories
- `word/`: Word application-specific settings and customizations

### `/docs/` - Documentation

Comprehensive project documentation:

```text
docs/
├── API.md                    # API documentation and function reference
├── CONTRIBUTORS.md           # Contributors list and acknowledgments  
├── CONTRIBUTING.md           # Guidelines for contributing
└── SECURITY.md              # Security policies and best practices
```

**Key Documents:**
- `API.md`: Complete API reference for developers
- `CONTRIBUTING.md`: How to contribute to the project
- `SECURITY.md`: Security policies for enterprise environments
- `CONTRIBUTORS.md`: List of contributors and their roles

### `/examples/` - Examples

Sample documents and test cases:

```text
examples/
└── prop-de-testes-01.docx   # Sample legislative document for testing
```

**Purpose:** Provide real-world examples and test documents for validation and demonstrations.

### `/scripts/` - Scripts

Installation and utility scripts:

```text
scripts/
├── install-chainsaw.ps1      # Main automated installer (PowerShell)
├── install-config.ini        # Installer configuration
└── INSTALL.md               # Detailed installation guide
```

**Key Scripts:**
- `install-chainsaw.ps1`: Automated installation with parameter support
- `install-config.ini`: Configuration for the installer
- `INSTALL.md`: Manual installation instructions

### `/src/` - Source Code

Main source code files:

```text
src/
└── chainsaw0.bas            # Main VBA module (~3000+ lines)
```

**Key Components in `chainsaw0.bas`:**
- Main processing function (`PadronizarDocumentoMain`)
- Configuration management system
- Backup and recovery system  
- Document validation functions
- Formatting and cleanup routines
- Logging and error handling
- Performance optimization code

### Root Files

#### `CHANGELOG.md`
Version history following [Keep a Changelog](https://keepachangelog.com/) format:
- **Added**: New features
- **Changed**: Changes in existing functionality  
- **Deprecated**: Soon-to-be removed features
- **Removed**: Now removed features
- **Fixed**: Bug fixes
- **Security**: Vulnerability fixes

#### `CODE_OF_CONDUCT.md`
Community guidelines based on [Contributor Covenant](https://www.contributor-covenant.org/) v2.1:
- Behavioral standards
- Enforcement procedures
- Contact information
- Scope definition

#### `LICENSE`
Apache 2.0 modified license with commercial restriction clause:
- Open source permissions
- Commercial use limitations
- Liability disclaimers
- Attribution requirements

#### `README.md`
Main project documentation including:
- Project overview and features
- Installation instructions
- Usage examples
- Configuration guide
- Contributing guidelines

## File Naming Conventions

### General Rules
- **Folders**: lowercase with hyphens (`folder-name`)
- **Documentation**: PascalCase for markdown (e.g., `README.md`, `CHANGELOG.md`)
- **Code Files**: lowercase with extensions (e.g., `chainsaw0.bas`)
- **Config Files**: lowercase with hyphens (e.g., `chainsaw-config.ini`)

### Language-Specific
- **VBA Files**: `.bas` extension
- **PowerShell Scripts**: `.ps1` extension  
- **Configuration**: `.ini` extension
- **Documentation**: `.md` extension

## Path Management

### Relative Paths
All internal references use relative paths from project root:
```text
./config/chainsaw-config.ini
./docs/SECURITY.md
./scripts/install-chainsaw.ps1
```

### Absolute Paths
Used only for system-specific locations:
- User Documents folder
- Word application directory
- Backup locations

## Access Patterns

### Read-Only Files
- Documentation (`/docs/`)
- Examples (`/examples/`)
- Source code (`/src/`)

### Configuration Files
- Read during startup
- Modified by installer
- User-customizable

### Generated Content
- Log files (temporary)
- Backup files (user documents)
- Processed documents

## Maintenance Guidelines

### Adding New Files
1. Follow naming conventions
2. Update this documentation
3. Add to appropriate `.gitignore` if needed
4. Update installation scripts if required

### Removing Files
1. Check for dependencies
2. Update documentation
3. Update installation/build scripts
4. Consider backward compatibility

### Reorganizing Structure  
1. Maintain backward compatibility
2. Update all path references
3. Test installation process
4. Update documentation comprehensively

## Integration Points

### Word Integration
- VBA module loaded into Word environment
- Configuration files accessed from user directories
- Assets referenced for document processing

### System Integration
- PowerShell installer modifies system settings
- Backup system interacts with file system
- Logging writes to user-accessible locations

### Development Integration
- Git repository structure
- Build/deployment processes
- Testing and validation workflows

---

**Last Updated:** 2025-09-25  
**Version:** 1.9.1-Alpha-8