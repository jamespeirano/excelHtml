# Adversys Excel Copilot for Excel

Excel add-in for AI-powered content generation!

## Technical Overview

### Architecture

- **Frontend:** Office.js Add-in (ES6+ JavaScript)
- **Authentication:** AWS Cognito with OAuth2.0
- **Backend:** AWS Lambda + API Gateway
- **Security:** JWT token-based authorization

### Core Features

- Asynchronous cell processing with exponential backoff
- Batch operation support with automatic retry mechanisms
- Real-time Excel range selection monitoring
- Stateful session management
- Graceful error handling and recovery

### Production Deployment

1. Install via Microsoft Office Store [link]
2. Authenticate with Adversys credentials
3. Grant required worksheet permissions


## License

Proprietary. Copyright © 2024 Adversys Corporation.
