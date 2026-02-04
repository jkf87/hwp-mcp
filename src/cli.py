# -*- coding: utf-8 -*-
import argparse
import json
import sys
import os

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    from src.tools.hwp_template_engine import HwpTemplateEngine
except ImportError:
    # Try relative import if running from src
    try:
        from tools.hwp_template_engine import HwpTemplateEngine
    except ImportError:
        print("Error: Could not import HwpTemplateEngine. Make sure you are in the project root or src directory.")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description="HWP Template Injection CLI Tool")
    parser.add_argument("template", help="Path to the HWP template file")
    parser.add_argument("data", help="Path to JSON data file or JSON string")
    parser.add_argument("output", help="Path to save the result HWP file")
    parser.add_argument("--visible", action="store_true", help="Show HWP window during processing")

    args = parser.parse_args()

    # Load data
    data = {}
    if os.path.exists(args.data):
        try:
            with open(args.data, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            print(f"Error reading JSON file: {e}")
            sys.exit(1)
    else:
        try:
            data = json.loads(args.data)
        except json.JSONDecodeError:
            print("Error: Data argument is neither a valid file path nor a valid JSON string.")
            sys.exit(1)

    print(f"Processing template: {args.template}")
    print(f"Output path: {args.output}")
    print(f"Data: {json.dumps(data, ensure_ascii=False, indent=2)}")

    try:
        engine = HwpTemplateEngine(visible=args.visible)
        
        if engine.process_template(args.template, data, args.output):
            print("Success! Document saved.")
            engine.quit()
        else:
            print("Failed to process template.")
            engine.quit()
            sys.exit(1)
            
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
