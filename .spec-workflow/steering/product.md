# Product Overview

## Product Purpose
MCP-Libre is a LibreOffice extension that provides Model Context Protocol (MCP) integration, enabling AI assistants to control and manipulate LibreOffice documents programmatically. It bridges the gap between AI tools and document processing by exposing LibreOffice's powerful UNO API through a standardized HTTP interface that MCP clients can consume.

The core problem it solves: AI assistants cannot natively interact with desktop office applications. MCP-Libre enables real-time document creation, editing, formatting, and export directly from AI conversations.

## Target Users
1. **AI/LLM Developers**: Building AI assistants that need document manipulation capabilities
2. **Knowledge Workers**: Using AI assistants (Claude, ChatGPT) for document-heavy workflows
3. **Automation Engineers**: Creating document processing pipelines with AI integration
4. **Research Teams**: Using AI to generate, edit, and manage research documents

Pain points addressed:
- Manual copy-paste between AI and documents
- Inability to see AI edits in real-time
- Limited programmatic access to LibreOffice from AI tools
- No standardized protocol for AI-document interaction

## Key Features

1. **Real-time Document Editing**: Insert, format, and modify text with instant visual feedback in LibreOffice
2. **HTTP API Integration**: REST API on localhost:8765 for external tool access
3. **Multi-format Support**: Create, save, and export to PDF, DOCX, ODT, HTML, TXT and more
4. **Comment/Annotation Support**: Add and retrieve document comments programmatically
5. **FastMCP Bridge**: Standard MCP protocol support for Claude Code and other MCP clients

## Business Objectives

- Enable seamless AI-to-document workflows without manual intervention
- Provide 10x faster document manipulation than external server approaches (direct UNO API)
- Establish MCP as the standard protocol for AI-document interaction
- Support the open-source office productivity ecosystem (LibreOffice)

## Success Metrics

- **Adoption**: Number of active MCP-Libre installations
- **Tool Coverage**: Percentage of common document operations supported
- **Latency**: Document operation response time < 100ms
- **Reliability**: 99%+ success rate for document operations

## Product Principles

1. **Direct Integration**: Embed directly in LibreOffice for maximum performance and minimal setup
2. **Protocol Standards**: Follow MCP specification for interoperability with any MCP client
3. **Real-time Feedback**: Users see document changes instantly as AI makes edits
4. **Surgical Precision**: Enable granular edits (paragraph-level, cursor positioning) not just bulk operations

## Monitoring & Visibility

- **Dashboard Type**: LibreOffice menu integration (Tools > MCP Server)
- **Real-time Updates**: HTTP server status visible in LibreOffice
- **Key Metrics Displayed**: Server running state, active connections
- **Sharing Capabilities**: HTTP API accessible to any local MCP client

## Future Vision

### Potential Enhancements
- **Enhanced Editing Tools**: Paragraph navigation, search/replace, cursor positioning for surgical edits
- **Document Outline**: View document structure without loading full content
- **Selection Operations**: Select and manipulate specific text ranges
- **Multi-document Coordination**: Work across multiple open documents simultaneously
- **Calc/Impress Support**: Extend surgical editing to spreadsheets and presentations
