# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.1] - 2025-03-28

### Added
- Option to include full message bodies when listing emails (`--include-bodies` flag)
- Option to filter out quoted content in email messages (`--hide-quoted` flag)
- Enhancement to `read-email` command to support hiding quoted content
- Support for these options in the MCP server implementation
- Improved email body handling with original content preservation

### Changed
- Enhanced `EmailDetails` interface to support original content and plain text conversions
- Updated `printEmailDetails` function to show information about removed quoted content

## [1.0.0] - 2025-03-26

### Added
- Initial release
- Authentication with Microsoft Graph API using client credentials flow
- List top-level mail folders
- List child folders of a specific mail folder
- Support for accessing mailboxes by specifying a user email (with proper permissions)
