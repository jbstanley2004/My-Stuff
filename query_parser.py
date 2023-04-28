import re

def parse_query(query):
    query = query.lower()
    if 'open' in query:
        match = re.search(r'open (.+)', query)
        if match:
            file_path = match.group(1)
            return f"Opening {file_path}..."
    elif 'read' in query:
        match = re.search(r'read (.+)', query)
        if match:
            file_path = match.group(1)
            return f"Reading {file_path}..."
    elif 'sum' in query:
        match = re.search(r'sum (.+)', query)
        if match:
            file_path = match.group(1)
            return f"Calculating sum for {file_path}..."
    elif 'average' in query:
        match = re.search(r'average (.+)', query)
        if match:
            file_path = match.group(1)
            return f"Calculating average for {file_path}..."
    elif 'count' in query:
        match = re.search(r'count (.+)', query)
        if match:
            file_path = match.group(1)
            return f"Counting rows in {file_path}..."
    else:
        return "I'm sorry, I didn't understand your query. Please try again with a valid query."

