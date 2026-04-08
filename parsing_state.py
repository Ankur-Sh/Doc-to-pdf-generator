class ParsingState:
    def __init__(self):
        self.questions = []
        self.images = []
        self.current_question = None
        self.current_section = None  # 'question', 'options', 'answer', 'explanation', 'source', 'image'
        
    def set_or_update_state(self, line: str):
        stripped = line.strip()
        
        # Skip empty lines
        if not stripped:
            return
        
        # Remove bold/italic markers for detection (but preserve in content)
        # This handles cases like "**Answer: c**" or "*Answer: a*"
        detection_line = stripped.replace("**", "").replace("*", "").strip()
            
        # Detect section markers (use detection_line for matching, but keep original formatting)
        if detection_line.lower().startswith("question:"):
            self._flush_current_question()
            self.current_question = {
                "question": "",
                "options": [],
                "answer": -1,
                "explanation": "",
                "source": "",
                "image": None
            }
            self.current_section = "question"
            content = stripped[9:].strip()  # Remove "question:" prefix
            if content:
                self.current_question["question"] += content + "\n"
        elif (detection_line.lower().startswith("option a:") or 
              detection_line.lower().startswith("a)") or
              (len(detection_line) > 1 and detection_line[0].lower() == 'a' and detection_line[1] in [')', '.', ':'])):
            self.current_section = "options"
            # Extract content preserving formatting
            if ":" in stripped:
                content = stripped.split(":", 1)[1].strip()
            elif stripped.lower().startswith("a)"):
                content = stripped[2:].strip()
            elif stripped.lower().startswith("a."):
                content = stripped[2:].strip()
            else:
                # Find where 'a' ends
                for i in range(1, min(4, len(stripped))):
                    if stripped[i] in [')', '.', ':']:
                        content = stripped[i+1:].strip()
                        break
                else:
                    content = stripped.strip()
            if content:
                self.current_question["options"].append(content)
        elif (detection_line.lower().startswith("option b:") or 
              detection_line.lower().startswith("b)") or
              (len(detection_line) > 1 and detection_line[0].lower() == 'b' and detection_line[1] in [')', '.', ':'])):
            self.current_section = "options"
            if ":" in stripped:
                content = stripped.split(":", 1)[1].strip()
            elif stripped.lower().startswith("b)"):
                content = stripped[2:].strip()
            elif stripped.lower().startswith("b."):
                content = stripped[2:].strip()
            else:
                for i in range(1, min(4, len(stripped))):
                    if stripped[i] in [')', '.', ':']:
                        content = stripped[i+1:].strip()
                        break
                else:
                    content = stripped.strip()
            if content:
                self.current_question["options"].append(content)
        elif (detection_line.lower().startswith("option c:") or 
              detection_line.lower().startswith("c)") or
              (len(detection_line) > 1 and detection_line[0].lower() == 'c' and detection_line[1] in [')', '.', ':'])):
            self.current_section = "options"
            if ":" in stripped:
                content = stripped.split(":", 1)[1].strip()
            elif stripped.lower().startswith("c)"):
                content = stripped[2:].strip()
            elif stripped.lower().startswith("c."):
                content = stripped[2:].strip()
            else:
                for i in range(1, min(4, len(stripped))):
                    if stripped[i] in [')', '.', ':']:
                        content = stripped[i+1:].strip()
                        break
                else:
                    content = stripped.strip()
            if content:
                self.current_question["options"].append(content)
        elif (detection_line.lower().startswith("option d:") or 
              detection_line.lower().startswith("d)") or
              (len(detection_line) > 1 and detection_line[0].lower() == 'd' and detection_line[1] in [')', '.', ':'])):
            self.current_section = "options"
            if ":" in stripped:
                content = stripped.split(":", 1)[1].strip()
            elif stripped.lower().startswith("d)"):
                content = stripped[2:].strip()
            elif stripped.lower().startswith("d."):
                content = stripped[2:].strip()
            else:
                for i in range(1, min(4, len(stripped))):
                    if stripped[i] in [')', '.', ':']:
                        content = stripped[i+1:].strip()
                        break
                else:
                    content = stripped.strip()
            if content:
                self.current_question["options"].append(content)
        elif (detection_line.lower().startswith("answer:") or 
              detection_line.lower().startswith("correct answer:") or
              detection_line.lower().startswith("solution:") or
              detection_line.lower().startswith("correct option:")):
            self.current_section = "answer"
            # Extract answer text after colon (remove formatting markers)
            if ":" in detection_line:
                answer_text = detection_line.split(":", 1)[1].strip()
            else:
                answer_text = detection_line.strip()
            
            # Remove any remaining bold/italic markers
            answer_text = answer_text.replace("**", "").replace("*", "").strip()
            
            answer_text_lower = answer_text.lower().strip()
            
            # First try: Match by option letter (a, b, c, d) or number (0, 1, 2, 3)
            if answer_text_lower.startswith("a") or answer_text_lower == "0" or answer_text_lower.startswith("(a)"):
                self.current_question["answer"] = 0
            elif answer_text_lower.startswith("b") or answer_text_lower == "1" or answer_text_lower.startswith("(b)"):
                self.current_question["answer"] = 1
            elif answer_text_lower.startswith("c") or answer_text_lower == "2" or answer_text_lower.startswith("(c)"):
                self.current_question["answer"] = 2
            elif answer_text_lower.startswith("d") or answer_text_lower == "3" or answer_text_lower.startswith("(d)"):
                self.current_question["answer"] = 3
            else:
                # Second try: Match by exact option text (case-insensitive)
                # This handles cases like "2 and 3 only", "All of the above", etc.
                if self.current_question and len(self.current_question["options"]) > 0:
                    for idx, option in enumerate(self.current_question["options"]):
                        # Compare normalized text (remove punctuation, case-insensitive)
                        option_normalized = option.lower().strip().rstrip('.')
                        answer_normalized = answer_text_lower.rstrip('.')
                        
                        # Exact match
                        if option_normalized == answer_normalized:
                            self.current_question["answer"] = idx
                            break
                        # Partial match (answer contains option text or vice versa)
                        elif option_normalized in answer_normalized or answer_normalized in option_normalized:
                            # Prefer longer match
                            if len(option_normalized) > 5:  # Only for substantial matches
                                self.current_question["answer"] = idx
                                break
        elif (detection_line.lower().startswith("explanation:") or 
              detection_line.lower().startswith("solution explanation:")):
            self.current_section = "explanation"
            if ":" in stripped:
                content = stripped.split(":", 1)[1].strip()
            else:
                content = stripped.strip()
            if content:
                self.current_question["explanation"] += content + "\n"
        elif (detection_line.lower().startswith("source:") or 
              detection_line.lower().startswith("reference:")):
            self.current_section = "source"
            if ":" in stripped:
                content = stripped.split(":", 1)[1].strip()
            else:
                content = stripped.strip()
            if content:
                self.current_question["source"] += content + "\n"
        elif detection_line.lower().startswith("image:") or stripped.startswith("[IMAGE:"):
            # Extract image index if it's a marker like [IMAGE:0]
            if stripped.startswith("[IMAGE:"):
                # Keep the marker in question text - will be processed during rendering
                # Always add [IMAGE:] markers to question text if we have a current question
                # Images can appear anywhere in the question content
                if self.current_question:
                    # If we're not in question section, switch to it
                    if self.current_section != "question":
                        self.current_section = "question"
                    self.current_question["question"] += line
            else:
                # Regular "image:" label - set section but don't add to question
                self.current_section = "image"
                # For now, just store empty image
                if self.current_question:
                    self.current_question["image"] = ""
        else:
            # Continue current section
            if self.current_question:
                if self.current_section == "question":
                    # Check if this line looks like it's starting a new section (options/answer/explanation)
                    # If so, don't add it to question text - it will be handled by the section detection
                    detection_lower = stripped.lower()
                    looks_like_option = (detection_lower.startswith(('a)', 'b)', 'c)', 'd)', 'option a', 'option b', 'option c', 'option d')) or
                                       (len(detection_lower) > 1 and detection_lower[0] in ['a', 'b', 'c', 'd'] and detection_lower[1] in [')', '.', ':']))
                    looks_like_answer = detection_lower.startswith(('answer:', 'correct answer:', 'solution:', 'correct option:'))
                    looks_like_explanation = detection_lower.startswith(('explanation:', 'solution explanation:'))
                    looks_like_source = detection_lower.startswith(('source:', 'reference:'))
                    
                    # If this line looks like it's starting a new section, don't add to question
                    if looks_like_option or looks_like_answer or looks_like_explanation or looks_like_source:
                        # This line will be processed in the next iteration when section changes
                        # Don't add it to question text
                        pass
                    else:
                        # Include table markdown lines in question text to preserve order
                        # Tables will be extracted later by parse_question_and_table
                        self.current_question["question"] += line
                elif self.current_section == "options":
                    # If we're in options section and line doesn't start with a/b/c/d, 
                    # it might be continuation of previous option
                    if len(self.current_question["options"]) > 0:
                        # Append to last option (handles multi-line options)
                        self.current_question["options"][-1] += " " + line.strip()
                elif self.current_section == "answer":
                    # Answer might be on next line, try to parse it
                    answer_text = stripped.lower().strip()
                    if answer_text and self.current_question["answer"] == -1:
                        # Try to match answer from this line
                        if answer_text.startswith("a") or answer_text == "0" or answer_text.startswith("(a)"):
                            self.current_question["answer"] = 0
                        elif answer_text.startswith("b") or answer_text == "1" or answer_text.startswith("(b)"):
                            self.current_question["answer"] = 1
                        elif answer_text.startswith("c") or answer_text == "2" or answer_text.startswith("(c)"):
                            self.current_question["answer"] = 2
                        elif answer_text.startswith("d") or answer_text == "3" or answer_text.startswith("(d)"):
                            self.current_question["answer"] = 3
                        else:
                            # Try matching with options (full text match)
                            for idx, option in enumerate(self.current_question["options"]):
                                option_normalized = option.lower().strip().rstrip('.')
                                answer_normalized = answer_text.rstrip('.')
                                if option_normalized == answer_normalized:
                                    self.current_question["answer"] = idx
                                    break
                elif self.current_section == "explanation":
                    self.current_question["explanation"] += line
                elif self.current_section == "source":
                    self.current_question["source"] += line
    
    def _flush_current_question(self):
        if self.current_question:
            # Clean up trailing newlines
            self.current_question["question"] = self.current_question["question"].strip()
            self.current_question["explanation"] = self.current_question["explanation"].strip()
            self.current_question["source"] = self.current_question["source"].strip()
            
            # If answer is still -1, try to infer it from explanation or other clues
            if self.current_question["answer"] == -1 and len(self.current_question["options"]) > 0:
                # Look for answer clues in explanation
                explanation_lower = self.current_question["explanation"].lower()
                for idx, option in enumerate(self.current_question["options"]):
                    option_lower = option.lower().strip()
                    # Check if explanation mentions the option as correct
                    if f"correct" in explanation_lower and option_lower in explanation_lower:
                        # Check if it's clearly indicating this option
                        if any(phrase in explanation_lower for phrase in [
                            f"option {chr(97+idx)}", f"({chr(97+idx)})", f"{chr(97+idx)} is correct",
                            f"answer is {chr(97+idx)}", f"correct answer is {option_lower}"
                        ]):
                            self.current_question["answer"] = idx
                            break
            
            self.questions.append(self.current_question)
            self.current_question = None
    
    def flush_state(self):
        """Finalize and add the last question if exists"""
        self._flush_current_question()


