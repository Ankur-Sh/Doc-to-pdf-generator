#!/usr/bin/env python3
"""Test table rendering"""
import sys
sys.path.insert(0, '.')

# Import without triggering main execution
import importlib.util
spec = importlib.util.spec_from_file_location("docx_gen", "docx_generator.py")
docx_gen = importlib.util.module_from_spec(spec)

# Mock input to avoid EOF error
import builtins
original_input = builtins.input
builtins.input = lambda x: "no"

try:
    spec.loader.exec_module(docx_gen)
finally:
    builtins.input = original_input

from docx_reader import extract_text_from_docx
from parsing_state import ParsingState

# Extract and parse
lines, imgs = extract_text_from_docx('files_to_convert/Geo- Practice test 1 (Final).docx')
parsing_state = ParsingState()
content = ''.join(lines).replace('***', '**')
lines = content.splitlines(keepends=True)
lines = [l for l in lines if not (len(set(l)) == 1 and l[0] == '\n')]
lines = [l.replace('\\', '') for l in lines]

for line in lines:
    parsing_state.set_or_update_state(line)

parsing_state.flush_state()

# Find question with table
for i, q in enumerate(parsing_state.questions):
    if '|' in q['question'] or 'Plate Boundary' in q['question']:
        print(f'\n=== Question {i+1} ===')
        print(f'Question text (first 300 chars): {q["question"][:300]}')
        print(f'\n--- Parsing table ---')
        before, rows, after = docx_gen.parse_question_and_table(q['question'])
        print(f'Before: "{before[:150]}"')
        print(f'Table rows found: {len(rows)}')
        for j, row in enumerate(rows[:5]):
            print(f'  Row {j}: {row}')
        print(f'After: "{after[:150]}"')
        break






