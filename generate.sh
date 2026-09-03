#!/bin/bash

# TravelGuard AI Presentation Generator
# This script generates a professional 10-slide PowerPoint presentation

echo "🚀 TravelGuard AI Presentation Generator"
echo "========================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3.7 or higher."
    exit 1
fi

# Install dependencies
echo "📦 Installing dependencies..."
pip install -r requirements.txt

if [ $? -ne 0 ]; then
    echo "❌ Failed to install dependencies."
    exit 1
fi

echo "✅ Dependencies installed successfully."
echo ""

# Generate presentation
echo "📊 Generating presentation..."
python3 create_presentation.py

if [ $? -ne 0 ]; then
    echo "❌ Failed to generate presentation."
    exit 1
fi

echo ""
echo "✅ Presentation generated successfully!"
echo "📁 File: TravelGuard_AI_Presentation.pptx"
echo ""
echo "🎯 Your 10-slide presentation is ready for the hackathon!"
echo ""
echo "Slides included:"
echo "  1. Cover - Title & Tagline"
echo "  2. Problem - Cascading disruptions"
echo "  3. Solution - Predict → Detect → Recover"
echo "  4. Architecture - 5-layer tech stack"
echo "  5. Intelligence - ML + Neo4j"
echo "  6. Detection - Real-time logic"
echo "  7. Recovery - Multi-criteria ranking"
echo "  8. Orchestration - AI + Live demo"
echo "  9. Innovations - 6 key differentiators"
echo "  10. Conclusion - Why we win"
echo ""
echo "✨ NO AI-GENERATED CONTENT - Hand-crafted presentation"
echo "🏆 Hackathon-compliant and ready to present!"
