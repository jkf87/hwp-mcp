# -*- coding: utf-8 -*-
import sys
import os
import re
import logging
from bs4 import BeautifulSoup, NavigableString

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__))))

from src.tools.hwp_controller import HwpController

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("generate-minutes")

def get_bullet_char(depth):
    """Returns a bullet character based on depth."""
    bullets = ["●", "○", "■", "□"]
    return bullets[depth % len(bullets)]

def process_node(node, level=0):
    """Recursively process nodes to extract text with indentation."""
    lines = []
    
    # Handle Text nodes
    if isinstance(node, NavigableString):
        text = node.strip()
        if text:
            return [text]
        return []

    # Skip some tags
    if node.name in ['style', 'script', 'meta', 'title', 'head']:
        return []

    # Determine indentation and prefix
    indent = "  " * level
    prefix = ""
    
    # Check if it's a list item
    is_list_item = node.name == 'li'
    if is_list_item:
        # Determine bullet based on depth relative to list nesting
        # We can pass 'list_depth' separately or estimate it.
        # Simple approach: use 'level' which we will increment for nested lists.
        # But 'li' is inside 'ul'. 'ul' increases level?
        pass

    # Process children
    # We need to separate "direct text" from "nested block children" for 'li'
    
    current_line_parts = []
    
    for child in node.children:
        if isinstance(child, NavigableString):
            text = child.strip()
            if text:
                current_line_parts.append(text)
        elif child.name in ['ul', 'ol']:
            # If we have accumulated text for this item, add it as a line first
            if current_line_parts:
                full_line = " ".join(current_line_parts)
                # Apply bullet if this was an li
                if is_list_item:
                     # Calculate bullet style based on level. 
                     # Level 0 is top. 
                     # If we are in li, we are at least inside one ul.
                     # Let's say level=0 is root.
                     # ul -> level+1. li -> uses level-1 (parent ul's level)?
                     # Let's adjust level logic.
                     b_char = get_bullet_char(max(0, level - 1))
                     lines.append(f"{indent}{b_char} {full_line}")
                else:
                     lines.append(f"{indent}{full_line}")
                current_line_parts = []
            
            # Recurse into list, increasing level
            lines.extend(process_node(child, level + 1))
            
        elif child.name in ['div', 'p', 'h1', 'h2', 'h3']:
             # Block elements
             # Flush current text
             if current_line_parts:
                full_line = " ".join(current_line_parts)
                if is_list_item:
                     b_char = get_bullet_char(max(0, level - 1))
                     lines.append(f"{indent}{b_char} {full_line}")
                else:
                     lines.append(f"{indent}{full_line}")
                current_line_parts = []
             
             # Recurse (divs might not increase level unless indented class?)
             # The HTML has <div class="indented">
             next_level = level
             if child.has_attr('class') and 'indented' in child['class']:
                 next_level += 1
                 
             lines.extend(process_node(child, next_level))
             
        elif child.name == 'li':
            # Direct li inside (shouldn't happen in valid HTML usually without UL, but possible)
            lines.extend(process_node(child, level)) # Li handles its own bulleting logic
            
        else:
            # Inline elements (strong, span, etc.) - treat as text
            # Recurse but treat results as inline text parts
            # This is tricky because process_node returns lines.
            # Simplified: just get text.
            child_text = child.get_text().strip()
            if child_text:
                current_line_parts.append(child_text)

    # Flush remaining text
    if current_line_parts:
        full_line = " ".join(current_line_parts)
        if is_list_item:
             b_char = get_bullet_char(max(0, level - 1))
             lines.append(f"{indent}{b_char} {full_line}")
        else:
             # If it's a p or header, usually it's a top level line
             lines.append(f"{indent}{full_line}")

    return lines

def extract_html_content_structured(html_path):
    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()
    
    soup = BeautifulSoup(html_content, "html.parser")
    title = soup.title.string if soup.title else "회의록"
    
    page_body = soup.find("div", class_="page-body")
    if not page_body:
        return title, soup.get_text()

    # Iterate over top-level children of page-body
    lines = []
    
    # We want to traverse logic.
    # page-body contains:
    # <div style="display:contents"> <p> ... </p> </div>
    # <div style="display:contents"> <ul class="bulleted-list"> ... </ul> </div>
    
    # We can iterate children of page_body
    # But often there are wrapper divs.
    
    # Let's simplify: find all 'p' (headers) and 'ul' (lists) in order?
    # No, order matters. Recursive traversal of page_body is safest.
    
    lines = process_node(page_body, level=0)
    
    # Filter empty lines
    lines = [line for line in lines if line.strip()]
    
    return title, "\n".join(lines)


def main():
    try:
        html_path = "260112 스토리보드 2차 검토 회의.html"
        template_path = "회의록 양식.hwp"
        output_path = "260112(회의록) 이에이트 웹 프론트엔드 스토리보드 2차 검토 회의 (2).hwpx"
        
        # 1. Extract content
        title, content = extract_html_content_structured(html_path)
        logger.info(f"Title: {title}")
        logger.info(f"Content length: {len(content)}")
        
        # 2. Open HWP
        hwp = HwpController()
        if not hwp.open_document(os.path.abspath(template_path)):
            logger.error("Failed to open template")
            return

        # 3. Replace placeholders
        # Title
        hwp.replace_text("[회의명]", title, replace_all=True)
        
        # Date
        date_str = "2026. 01. 12."
        hwp.replace_text("yyyy. mm. dd.", date_str, replace_all=True)
        
        # Place
        place = "이에이트"
        hwp.replace_text("[회의장소]", place, replace_all=True)
        
        # 4. Content
        if hwp.find_text("[회의록 내용]"):
            hwp.hwp.Run("Delete")
            hwp.insert_text(content)
        else:
            logger.warning("[회의록 내용] placeholder not found")
            
        # 5. Save
        if hwp.save_document(os.path.abspath(output_path)):
            logger.info(f"Saved to {output_path}")
        else:
            logger.error("Failed to save document")
            
        hwp.close_document()
        hwp.disconnect()
        
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()