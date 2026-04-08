#!/usr/bin/env python3
"""
Debug script to inspect DOCX file structure and identify answer format
"""
import sys
from docx_reader import extract_text_from_docx

if len(sys.argv) < 2:
    print("Usage: python3 debug_docx.py <docx_file_path>")
    sys.exit(1)

file_path = sys.argv[1]

print("=" * 60)
print(f"Analyzing: {file_path}")
print("=" * 60)
print()

try:
    lines, images = extract_text_from_docx(file_path)
    
    print(f"Total lines extracted: {len(lines)}")
    print(f"Images found: {len(images)}")
    print()
    print("First 50 lines of extracted content:")
    print("-" * 60)
    for i, line in enumerate(lines[:50], 1):
        print(f"{i:3}: {line.rstrip()}")
    
    print()
    print("=" * 60)
    print("Looking for answer patterns...")
    print("=" * 60)
    
    # Look for answer patterns
    answer_patterns = []
    for i, line in enumerate(lines):
        line_lower = line.lower().strip()
        if any(keyword in line_lower for keyword in ['answer', 'solution', 'correct']):
            answer_patterns.append((i+1, line.strip()))
    
    if answer_patterns:
        print(f"\nFound {len(answer_patterns)} potential answer lines:")
        for line_num, line_text in answer_patterns[:10]:  # Show first 10
            print(f"  Line {line_num}: {line_text[:80]}")
    else:
        print("\n⚠ No answer patterns found!")
        print("  The file might not have explicit 'Answer:' labels.")
        print("  Answers might be embedded in the text or in a different format.")
    
    print()
    print("=" * 60)
    print("Sample question structure:")
    print("=" * 60)
    
    # Find question blocks
    in_question = False
    question_count = 0
    for i, line in enumerate(lines):
        if 'question' in line.lower() and ':' in line:
            if question_count < 3:  # Show first 3 questions
                print(f"\n--- Question {question_count + 1} (around line {i+1}) ---")
                # Show next 15 lines
                for j in range(i, min(i+15, len(lines))):
                    print(f"{j+1:3}: {lines[j].rstrip()}")
                question_count += 1
            in_question = True
        if question_count >= 3:
            break

except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()






