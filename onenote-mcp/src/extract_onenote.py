import sys
import os
import json
import tempfile
import subprocess
import shutil

def install_aspose_note():
    try:
        from aspose.note import Document, Page, RichText
    except ImportError:
        # Install aspose-note
        print("Installing aspose-note Python package...", file=sys.stderr)
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "aspose-note"])
        except Exception as e:
            print(f"Error installing aspose-note: {e}", file=sys.stderr)
            sys.exit(1)

def extract_one_file(file_path):
    install_aspose_note()
    from aspose.note import Document, Page, RichText

    try:
        doc = Document(file_path)
    except Exception as e:
        print(f"Error loading document {file_path}: {e}", file=sys.stderr)
        return []

    pages = []
    try:
        pages_nodes = doc.GetChildNodes(Page)
    except Exception as e:
        print(f"Error getting child nodes from {file_path}: {e}", file=sys.stderr)
        return []

    for page in pages_nodes:
        # Resolve page title
        title = "Untitled Page"
        if page.Title and page.Title.TitleText:
            title = page.Title.TitleText.Text or "Untitled Page"
        
        # Resolve text content
        content_html = ""
        try:
            rt_nodes = page.GetChildNodes(RichText)
            for rt in rt_nodes:
                text = rt.Text
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        cleaned = line.strip()
                        if cleaned:
                            # Simple HTML rendering
                            content_html += f"<p>{cleaned}</p>"
        except Exception as e:
            print(f"Error reading page content: {e}", file=sys.stderr)

        if not content_html:
            content_html = "<p>Empty Page</p>"

        pages.append({
            "title": title.strip(),
            "content": content_html
        })
    return pages

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_onenote.py <file_path>", file=sys.stderr)
        sys.exit(1)

    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}", file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.onepkg':
        # Create temp folder
        temp_dir = tempfile.mkdtemp(prefix="onenote_unpack_")
        try:
            # Expand CAB (.onepkg) using native Windows expand
            print(f"Unpacking .onepkg file into {temp_dir}...", file=sys.stderr)
            # expand -F:* <src> <dest>
            subprocess.run(["expand", "-F:*", file_path, temp_dir], check=True, stdout=subprocess.DEVNULL)

            # Find all .one files in the unpacked folder
            all_pages = []
            for root, dirs, files in os.walk(temp_dir):
                for f in files:
                    if f.lower().endswith('.one'):
                        full_p = os.path.join(root, f)
                        print(f"Processing unpacked file: {f}", file=sys.stderr)
                        pages = extract_one_file(full_p)
                        # Prefix page title with section name if appropriate
                        section_name = os.path.splitext(f)[0]
                        for page in pages:
                            page["title"] = f"{section_name} - {page['title']}"
                        all_pages.extend(pages)
            
            # Output result
            print(json.dumps(all_pages, indent=2))
        except Exception as e:
            print(f"Error extracting cabinet .onepkg file: {e}", file=sys.stderr)
            sys.exit(1)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    elif ext in ['.one', '.onex']:
        pages = extract_one_file(file_path)
        print(json.dumps(pages, indent=2))
    
    else:
        # Fallback to reading as standard file
        print(f"Unknown extension {ext}, reading as plain text", file=sys.stderr)
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            title = os.path.splitext(os.path.basename(file_path))[0]
            # simple html wrapper
            html = "".join([f"<p>{line.strip()}</p>" for line in content.split('\n') if line.strip()])
            print(json.dumps([{
                "title": title,
                "content": html or "<p>Empty note</p>"
            }], indent=2))
        except Exception as e:
            print(f"Error reading plain file: {e}", file=sys.stderr)
            sys.exit(1)

if __name__ == "__main__":
    main()
