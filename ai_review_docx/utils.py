from diff_match_patch import diff_match_patch

# ANSI color codes
ANSI_RED = "\033[91m"
ANSI_GREEN = "\033[92m"
ANSI_RESET = "\033[0m"


def reviewed_path(original_path: str) -> str:
    """
    Generates the file path for the reviewed document.
    """
    return original_path.replace(".docx", "_reviewed.docx")


def colored_console_diff(text1: str, text2: str) -> str:
    """
    Computes a diff and formats it with console color tags.
    """
    dmp = diff_match_patch()
    diffs = dmp.diff_main(text1, text2)
    dmp.diff_cleanupSemantic(diffs)
    
    colored_parts = []
    for (op, data) in diffs:
        if op == dmp.DIFF_DELETE:
            colored_parts.append(f"{ANSI_RED}{data}{ANSI_RESET}")
        elif op == dmp.DIFF_INSERT:
            colored_parts.append(f"{ANSI_GREEN}{data}{ANSI_RESET}")
        elif op == dmp.DIFF_EQUAL:
            colored_parts.append(data)
    return "".join(colored_parts)


def preview(text: str) -> str:
    """
    Returns a preview of the text.
    """
    if len(text) > 50:
        return text[:50] + "..."
    else:
        return text


def formatted_diff_for_docx(text1: str, text2: str) -> list[tuple[str, dict]]:
    """
    Computes a diff and returns formatted runs for Word documents.

    Returns:
        List of tuples (text, formatting_dict) where:
        - Deletions are formatted with red color and strikethrough
        - Insertions are formatted with green color
        - Unchanged text has no formatting
    """
    dmp = diff_match_patch()
    diffs = dmp.diff_main(text1, text2)
    dmp.diff_cleanupSemantic(diffs)

    formatted_runs = []
    for (op, data) in diffs:
        if op == dmp.DIFF_DELETE:
            # Red color with strikethrough for deletions
            formatted_runs.append((data, {"color": "FF0000", "strike": True}))
        elif op == dmp.DIFF_INSERT:
            # Green color for insertions
            formatted_runs.append((data, {"color": "00B050"}))
        elif op == dmp.DIFF_EQUAL:
            # No formatting for unchanged text
            formatted_runs.append((data, {}))

    return formatted_runs
