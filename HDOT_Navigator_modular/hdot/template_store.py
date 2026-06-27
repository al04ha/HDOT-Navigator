"""Stub template_store module."""

def list_docx_templates():
    return []

def find_matching_template(doc_type):
    return None

def extract_template_formatting(template_bytes):
    return ""

def save_docx_template(name, bytes_data):
    pass

def load_docx_template(name):
    return None

def delete_docx_template(name):
    pass

def save_template_override(doc_type, bytes_data):
    pass

def delete_template_override(doc_type):
    pass

def has_template_override(doc_type):
    return False

def _template_path(name):
    return None

def load_kb_folder_links():
    return {}

def save_kb_folder_links(links):
    pass

def get_kb_folder_for_doc_type(doc_type):
    return None

def get_kb_context_for_folder(folder):
    return ""

def get_kb_folders():
    return []
