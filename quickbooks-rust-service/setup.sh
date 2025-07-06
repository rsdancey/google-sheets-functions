#!/bin/bash

# QuickBooks to Google Sheets Setup Script

echo "🔧 Setting up QuickBooks to Google Sheets Integration"
echo "======================================================"

# Check if we're in the right directory
if [ ! -f "Cargo.toml" ]; then
    echo "❌ Error: Please run this script from the quickbooks-rust-service directory"
    exit 1
fi

# Step 1: Setup configuration
echo "📋 Step 1: Setting up configuration..."
if [ ! -f "config/config.toml" ]; then
    cp config/config.example.toml config/config.toml
    echo "✅ Created config/config.toml from example"
    echo "⚠️  Please edit config/config.toml with your settings"
else
    echo "✅ Configuration file already exists"
fi

# Step 2: Build the service
echo "📋 Step 2: Building the Rust service..."
cargo build --release

if [ $? -eq 0 ]; then
    echo "✅ Rust service built successfully"
else
    echo "❌ Failed to build Rust service"
    exit 1
fi

# Step 3: Check Google Apps Script setup
echo "📋 Step 3: Google Apps Script setup checklist..."
echo "   □ 1. Run setupQuickBooksIntegration() in your Google Apps Script"
echo "   □ 2. Copy the generated API key to config/config.toml"
echo "   □ 3. Deploy your Google Apps Script as a Web App"
echo "   □ 4. Copy the Web App URL to config/config.toml"
echo "   □ 5. Set permissions to 'Anyone' for the Web App"

echo ""
echo "🎉 Setup Complete!"
echo "======================================================"
echo "Next steps:"
echo "1. Edit config/config.toml with your specific settings"
echo "2. Ensure QuickBooks Desktop Enterprise is running"
echo "3. Run the service: cargo run --release"
echo ""
echo "📚 For detailed instructions, see README.md"
