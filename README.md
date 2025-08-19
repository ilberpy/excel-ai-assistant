# Excel AI Assistant

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Office Add-in](https://img.shields.io/badge/Office%20Add--in-Excel-blue.svg)](https://docs.microsoft.com/en-us/office/dev/add-ins/)

**The first open-source local AI-powered Excel assistant tool**

A powerful Excel add-in that integrates with local AI models (LM Studio) to provide intelligent data analysis, automated Excel operations, and natural language processing capabilities - all while keeping your data completely local and secure.

## 🌟 Features

### 🤖 AI Integration
- **Local AI Model Support**: Integration with LM Studio for complete data privacy
- **Real-time Streaming Responses**: ChatGPT-like streaming responses
- **Multi-Model Selection**: Choose from available AI models
- **Intelligent Data Analysis**: Automatic data analysis and Excel operations

### 📊 Excel Operations
- **Automatic Chart Creation**: Smart chart selection based on data
- **Smart Formatting**: Automatic table formatting and color coding
- **Data Filtering**: Advanced filtering and sorting capabilities
- **Calculations**: Automatic sum, average, and statistical calculations
- **Trend Analysis**: Automatic data trend detection
- **Anomaly Detection**: Find data abnormalities automatically

### 💬 Chat Interface
- **Modern Dark Theme**: Cursor-like dark theme design
- **Chat History**: Persistent chat records with load/delete functionality
- **Auto-scrolling**: Automatic screen scrolling while AI is typing
- **Response Control**: Stop AI responses at any time
- **Voice and Visual Input**: Voice commands and image upload support

### ⚙️ Settings & Customization
- **Theme Selection**: Dark, light, and blue theme options
- **Excel Settings**: Chart types, sizes, automatic formatting preferences
- **AI Model Management**: Model connection testing and updates
- **User Preferences**: Personalized settings and configurations

## 🚀 Installation

### Requirements
- Microsoft Excel (Desktop or Online)
- LM Studio (Local AI model server)
- Modern web browser

### Steps

1. **Install LM Studio**
   ```bash
   # Download and install LM Studio
   # https://lmstudio.ai/
   ```

2. **Project Setup**
   ```bash
   git clone https://github.com/ilberpy/excel-ai-assistant.git
   cd excel-ai-assistant
   npm install
   ```

3. **AI Model Configuration**
   ```bash
   # Load models in LM Studio
   # Start API server (port 1234)
   ```

4. **Excel Add-in Installation**
   ```bash
   npm run start
   # In Excel: Developer > Add-ins > Upload My Add-in
   ```

## 🔧 Configuration

### LM Studio Connection
```javascript
// ai_client.js
const baseUrl = 'http://192.168.1.5:1234'; // Enter your own IP address
```

### Excel Settings
```javascript
// Excel settings are stored in localStorage
{
  "chartType": "ColumnClustered",
  "autoFormatting": true,
  "zebraRows": true,
  "headerHighlighting": true
}
```

## 📖 Usage

### Basic Commands
- **"Analyze this data"** - Intelligent analysis of selected data
- **"Create a chart"** - Automatic chart generation
- **"Calculate totals"** - Column sum calculations
- **"Filter data"** - Smart data filtering
- **"Format table"** - Automatic table formatting

### Advanced Features
- **Voice Commands**: Give commands via microphone
- **Image Analysis**: Upload images for AI analysis
- **Chat History**: Reopen previous conversations
- **Theme Customization**: Personal theme selection

## 🏗️ Architecture

```
excel-ai-assistant/
├── app.js              # Main application logic
├── ai_client.js        # AI API client
├── index.html          # User interface
├── styles.css          # Style definitions
├── manifest.xml        # Office Add-in manifest
├── package.json        # Project dependencies
└── README.md           # Project documentation
```

## 🤝 Contributing

This project is open source and we welcome your contributions!

1. Fork the project
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Microsoft Office Add-ins team
- LM Studio developers
- Open source community
- All contributors

## 📞 Contact

- **GitHub Issues**: [Project Issues](https://github.com/ilberpy/excel-ai-assistant/issues)
- **Discussions**: [GitHub Discussions](https://github.com/ilberpy/excel-ai-assistant/discussions)

---

**⭐ If you like this project, don't forget to give it a star!**
