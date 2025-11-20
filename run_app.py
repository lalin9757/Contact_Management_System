#!/usr/bin/env python3
"""
Contact Management System Launcher
"""

import os
import sys
import traceback

def main():
    print("🚀 Starting Contact Management System...")
    print("📁 Project Path: E:\\5th Semester\\Contact Management System")
    
    try:
        # Check if required packages are installed
        try:
            import customtkinter
            import PIL
            import sqlite3
            print("✅ All dependencies are available")
        except ImportError as e:
            print(f"❌ Missing dependency: {e}")
            print("Please install required packages using:")
            print("pip install -r requirements.txt")
            input("Press Enter to exit...")
            return
        
        # Import and run the application
        from main import ContactManagementSystem
        
        print("✅ Application loaded successfully")
        print("🖥️  Launching GUI...")
        
        app = ContactManagementSystem()
        app.run()
        
    except Exception as e:
        print(f"❌ Error starting application: {e}")
        print("\n📋 Detailed error information:")
        traceback.print_exc()
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()