import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import sys
import os
import re
import argparse
import shutil
from datetime import datetime
import openpyxl

# ==========================================
# CSS STYLES
# ==========================================

CSS_SIGNAL = """
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, Cantarell, "Open Sans", "Helvetica Neue", sans-serif; background-color: #f7f7f7; margin: 0; padding: 0; color: #1b1b1b; }
.container { display: flex; height: 100vh; }
.sidebar { width: 320px; flex-shrink: 0; background-color: white; border-right: 1px solid #e6e6e6; overflow-y: auto; }
.main { flex: 1; display: flex; flex-direction: column; background-color: #ffffff; }
.chat-messages { flex: 1; overflow-y: auto; padding: 20px 40px; display: flex; flex-direction: column; }
.message { max-width: 60%; margin-bottom: 8px; padding: 10px 14px; border-radius: 16px; position: relative; font-size: 15px; line-height: 1.4; word-wrap: break-word; }
.message.received { align-self: flex-start; background-color: #f2f2f2; color: #1b1b1b; border-bottom-left-radius: 4px; margin-right: auto; }
.message.sent { align-self: flex-end; background-color: #2c6bed; color: white; border-bottom-right-radius: 4px; margin-left: auto; }
.sender-name { font-size: 12px; font-weight: bold; margin-bottom: 2px; color: #666; }
.message.sent .sender-name { display: none; }
.timestamp { font-size: 10px; opacity: 0.6; float: right; margin-left: 8px; margin-top: 4px; }
.message.sent .timestamp { color: rgba(255,255,255,0.8); }
.message.has-attachment { padding: 5px; }
.message.has-attachment .attachment img { max-width: 200px; border-radius: 6px; margin: 0; }

/* SIDEBAR STYLES */
.signal-sidebar-header { font-size: 20px; font-weight: bold; padding: 24px 20px 16px 20px; border-bottom: 1px solid #e6e6e6; color: #1b1b1b; }
.chat-item { height: 72px; padding: 0 20px; border-bottom: 1px solid #f5f5f5; cursor: pointer; display: flex; flex-direction: row; align-items: center; transition: background-color 0.1s; }
.chat-item:hover { background-color: #f5f5f5; }
.chat-item.active { background-color: #d1e1fb; border-left: 4px solid #2c6bed; }
.sidebar-avatar { width: 52px; height: 52px; font-size: 20px; margin-right: 16px; margin-bottom: 0; display: flex; align-items: center; justify-content: center; background-color: #ddd; border-radius: 50%; color: #333; font-weight: bold; }
.sidebar-name { font-weight: 600; font-size: 15px; color: #1b1b1b; }
.sidebar-info { font-size: 13px; color: #666; margin-top: 3px; }

/* SHARED / OTHER */
.chat-info-header { background-color: #f7f7f7; padding: 15px; border-bottom: 1px solid #ddd; display: flex; justify-content: center; align-items: center; gap: 40px; margin: 10px 20px 20px 20px; border-radius: 8px; }
.participant-avatar { width: 60px; height: 60px; border-radius: 50%; background-color: #ddd; color: #333; display: flex; align-items: center; justify-content: center; font-size: 24px; font-weight: bold; margin-bottom: 8px; }
.participant-card { display: flex; flex-direction: column; align-items: center; text-align: center; }
.participant-name { font-weight: 700; font-size: 16px; color: #1b1b1b; }
.participant-number { font-size: 13px; color: #666; margin-top: 2px; }
.translation { margin-top: 8px; padding: 8px 10px; border-left: 3px solid #fecb00; font-style: normal; font-size: 13.5px; background-color: rgba(255,255,255,0.95); color: #333; border-radius: 6px; display: block; clear: both; margin-bottom: 2px; }
.date-divider { align-self: center; color: #555; font-size: 13px; margin: 15px 0; font-weight: 500; text-align: center; }
.message.sent .translation { background-color: rgba(255,255,255,0.15); color: white; border-left-color: rgba(255,255,255,0.5); }
.attachment img { max-width: 200px; border-radius: 12px; margin-top: 5px; cursor: pointer; }
.attachment video { max-width: 250px; border-radius: 12px; margin-top: 5px; }
.attachment audio { max-width: 250px; margin-top: 5px; }
.search-box { padding: 10px 15px; background-color: #f7f7f7; border-bottom: 1px solid #e6e6e6; }
.search-input { width: 100%; padding: 8px 12px; border-radius: 6px; border: 1px solid #ddd; background-color: #e9e9e9; font-size: 14px; box-sizing: border-box; outline: none; }
.search-input:focus { background-color: white; border-color: #2c6bed; }

/* ADVANCED SEARCH - PREMIUM */
.search-results-container { display: none; flex: 1; overflow-y: auto; background-color: #fff; }
.search-header { padding: 16px 20px; font-weight: 700; font-size: 14px; color: #1b1b1b; background: #f7f7f7; border-bottom: 1px solid #e6e6e6; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; z-index: 5; }
.close-search { cursor: pointer; color: #2c6bed; font-weight: 600; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; padding: 4px 8px; border-radius: 4px; transition: background 0.2s; }
.close-search:hover { background-color: rgba(44, 107, 237, 0.1); }
.search-result-item { padding: 12px 20px; border-bottom: 1px solid #f5f5f5; cursor: pointer; transition: background 0.2s; display: flex; align-items: center; }
.search-result-item:hover { background-color: #f0f2f5; }
.search-avatar { width: 42px; height: 42px; border-radius: 50%; background-color: #e0e0e0; color: #555; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 16px; margin-right: 14px; flex-shrink: 0; }
.search-content { flex: 1; min-width: 0; }
.search-result-sender { font-size: 14px; font-weight: 600; color: #1b1b1b; margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.search-result-preview { font-size: 13px; color: #666; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; line-height: 1.4; }
.highlight-term { background-color: rgba(255, 235, 59, 0.5); border-radius: 2px; color: #000; font-weight: 500; padding: 0 2px; }
.highlight-msg { position: relative; animation: flashMsg 2s ease-out; }
.highlight-msg::after { content: ""; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background-color: rgba(255, 235, 59, 0.2); border-radius: inherit; pointer-events: none; opacity: 0; animation: fadeOverlay 3s forwards; }
@keyframes flashMsg { 0% { transform: scale(1.02); box-shadow: 0 4px 12px rgba(0,0,0,0.1); z-index: 10; } 100% { transform: scale(1); box-shadow: none; z-index: 1; } }
@keyframes fadeOverlay { 0% { opacity: 1; } 100% { opacity: 0; } }
"""

CSS_WHATSAPP = """
body { font-family: "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #d1d7db; margin: 0; padding: 0; color: #111b21; }
.container { display: flex; height: 100vh; }
.sidebar { width: 350px; flex-shrink: 0; background-color: white; border-right: 1px solid #e9edef; overflow-y: auto; }
.main { flex: 1; display: flex; flex-direction: column; background-color: #efe7dd; background-image: url("https://user-images.githubusercontent.com/15075759/28719144-86dc0f70-73b1-11e7-911d-60d70fcded21.png"); }
.chat-messages { flex: 1; overflow-y: auto; padding: 20px 60px; display: flex; flex-direction: column; }
.message { max-width: 65%; margin-bottom: 8px; padding: 6px 7px 8px 9px; border-radius: 7.5px; position: relative; font-size: 14.2px; line-height: 19px; box-shadow: 0 1px 0.5px rgba(11,20,26,0.13); word-wrap: break-word; }
.message.received { align-self: flex-start; background-color: #ffffff; color: #111b21; border-top-left-radius: 0; margin-right: auto; }
.message.sent { align-self: flex-end; background-color: #d9fdd3; color: #111b21; border-top-right-radius: 0; margin-left: auto; }
.timestamp { font-size: 11px; float: right; margin-left: 10px; margin-top: 5px; color: #667781; }
.chat-item { height: 72px; display: flex; align-items: center; padding: 0 15px; border-bottom: 1px solid #f0f2f5; cursor: pointer; flex-direction: row; }
.chat-item:hover { background-color: #f5f6f6; }
.chat-item.active { background-color: #ebebeb; border-left: 4px solid #00a884; }
.chat-info-header { background-color: rgba(255,255,255,0.92); padding: 15px; border-radius: 12px; margin: 20px auto; max-width: 600px; display: flex; justify-content: center; align-items: center; gap: 30px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
.chat-header-bar { background-color: #f0f2f5; padding: 10px 16px; border-left: 1px solid #d1d7db; display: flex; align-items: center; height: 60px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); z-index: 10; }
.participant-avatar { width: 60px; height: 60px; border-radius: 50%; background-color: #e9edef; color: #fff; display: flex; align-items: center; justify-content: center; font-size: 24px; font-weight: bold; margin-bottom: 8px; border: 3px solid white; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
.participant-avatar.owner { background-color: #00a884; }
.participant-avatar.contact { background-color: #555; }
.participant-card { display: flex; flex-direction: column; align-items: center; text-align: center; }
.participant-name { font-weight: 600; font-size: 15px; color: #111b21; }
.participant-number { font-size: 12px; color: #667781; }
.sidebar-avatar { width: 45px; height: 45px; font-size: 18px; margin-right: 15px; margin-bottom: 0; display: flex; align-items: center; justify-content: center; background-color: #e9edef; border-radius: 50%; color: #fff; font-weight: bold; }
.sidebar-name { font-weight: 600; font-size: 15px; color: #111b21; }
.sidebar-info { font-size: 13px; color: #667781; }
.translation { margin-top: 6px; padding: 6px 8px; border-left: 3px solid #ffcc00; font-size: 13px; background-color: rgba(0,0,0,0.03); color: #333; border-radius: 4px; font-style: italic; }
.date-divider { text-align: center; align-self: center; margin: 10px 0; padding: 5px 12px; background-color: rgba(255,255,255,0.9); border-radius: 7.5px; box-shadow: 0 1px 0.5px rgba(11,20,26,0.13); color: #555; font-size: 12.5px; }
.attachment img { max-width: 200px; border-radius: 6px; margin-top: 4px; cursor: pointer; }
.attachment video { max-width: 200px; border-radius: 8px; margin-top: 5px; }
.attachment audio { max-width: 100%; margin-top: 5px; }
.message.received .sender-name { font-weight: bold; font-size: 13.5px; color: #d64937; margin-bottom: 4px; display: block; }
.message.sent .sender-name { font-weight: bold; font-size: 13.5px; color: #555; margin-bottom: 4px; display: block; opacity: 0.9; }
.message.has-attachment { padding: 5px; }
.message.has-attachment .attachment img { max-width: 200px; border-radius: 6px; margin: 0; }
.search-box { padding: 10px; background-color: #f0f2f5; border-bottom: 1px solid #d1d7db; }
.search-input { width: 100%; padding: 7px 12px; border-radius: 8px; border: none; background-color: #ffffff; font-size: 14px; box-sizing: border-box; height: 35px; outline: none; }

/* ADVANCED SEARCH - PREMIUM */
.search-results-container { display: none; flex: 1; overflow-y: auto; background-color: #fff; }
.search-header { padding: 16px 20px; font-weight: 700; font-size: 14px; color: #1b1b1b; background: #f7f7f7; border-bottom: 1px solid #e6e6e6; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; z-index: 5; }
.close-search { cursor: pointer; color: #2c6bed; font-weight: 600; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; padding: 4px 8px; border-radius: 4px; transition: background 0.2s; }
.close-search:hover { background-color: rgba(44, 107, 237, 0.1); }
.search-result-item { padding: 12px 20px; border-bottom: 1px solid #f5f5f5; cursor: pointer; transition: background 0.2s; display: flex; align-items: center; }
.search-result-item:hover { background-color: #f0f2f5; }
.search-avatar { width: 42px; height: 42px; border-radius: 50%; background-color: #e0e0e0; color: #555; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 16px; margin-right: 14px; flex-shrink: 0; }
.search-content { flex: 1; min-width: 0; }
.search-result-sender { font-size: 14px; font-weight: 600; color: #1b1b1b; margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.search-result-preview { font-size: 13px; color: #666; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; line-height: 1.4; }
.highlight-term { background-color: rgba(255, 235, 59, 0.5); border-radius: 2px; color: #000; font-weight: 500; padding: 0 2px; }
.highlight-msg { position: relative; animation: flashMsg 2s ease-out; }
.highlight-msg::after { content: ""; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background-color: rgba(255, 235, 59, 0.2); border-radius: inherit; pointer-events: none; opacity: 0; animation: fadeOverlay 3s forwards; }
@keyframes flashMsg { 0% { transform: scale(1.02); box-shadow: 0 4px 12px rgba(0,0,0,0.1); z-index: 10; } 100% { transform: scale(1); box-shadow: none; z-index: 1; } }
@keyframes fadeOverlay { 0% { opacity: 1; } 100% { opacity: 0; } }
"""

# ==========================================
# HELPERS
# ==========================================

def clean_text(text):
    if text is None:
        return ""
    
    text = str(text)
    if not text or text.lower() == 'nan':
         return ""

    patterns = [
        r'Etichette:.*?(?=\n|$)',
        r'Creato:.*?(?=\n|$)',
        r'Modificato:.*?(?=\n|$)',
        r'Descrizione:.*?(?=\n|$)',
        r'Generator:.*?(?=\n|$)',
    ]
    cleaned = text
    for p in patterns:
        cleaned = re.sub(p, '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\n\s*\n', '\n', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\n\s*\n', '\n', cleaned).strip()
    cleaned = cleaned.replace("_x000d_", "")
    return cleaned

def parse_participants_intelligent(data, chat_id=""):
    """
    Intelligent logic to extract Owner and Contact from chat data.
    Returns: (owner_name, owner_number, contact_name, contact_number)
    """
    # 1. Get Owner from metadata if available (populated by WhatsAppParser)
    owner_name = data.get('owner', 'Unknown')
    owner_num = "" 
    if owner_name == "Unknown" or not owner_name:
         owner_name = "Proprietario"

    # Infer Owner from Messages if not known
    if owner_name == "Proprietario" or owner_name == "Tu":
        # Scan messages for is_sent = True
        msgs = data.get('messages', [])
        for m in msgs:
            if m.get('is_sent', False):
                possible_owner = m.get('sender')
                if possible_owner and possible_owner != "Tu" and possible_owner != "Proprietario":
                     owner_name = possible_owner
                     break
    
    # 2. Get Contact Name from participants field (already cleaned by Parser)
    contact_name = data.get('participants', '')
    contact_num = ""

    # If contact_name still contains "Sconosciuto" and we can't do better, keep it.
    if not contact_name or contact_name == "Unknown":
         contact_name = str(chat_id).replace("Chat ", "")

    # Fallback/Legacy Logic for Cellebrite Parser
    if "_x000d_" in contact_name or "\n" in contact_name:
        parts_raw = contact_name
        segments = re.split(r'_x000d_|\n', parts_raw)
        c_names = []
        for seg in segments:
            seg = seg.strip()
            if not seg: continue
            seg = re.sub(r'\(proprietario\)', '', seg, flags=re.IGNORECASE)
            seg = re.sub(r'\(owner\)', '', seg, flags=re.IGNORECASE).strip()
            if seg and "sconosciuto" not in seg.lower():
                 c_names.append(seg)
        if c_names:
             contact_name = " & ".join(c_names)

    # Clean Owner from Contact Name if present
    if owner_name and owner_name != "Proprietario" and owner_name != "Tu":
        if owner_name in contact_name and " & " in contact_name:
             # Basic replace: Remove "Alice & " or " & Alice"
             contact_name = contact_name.replace(f"{owner_name} & ", "").replace(f" & {owner_name}", "")
    
    # Final cleanup: Separate Number from Name if mixed "12345 Name"
    # This fixes "+39347... Name" appearing in sidebar
    if contact_name:
        # Regex for long number sequence
        match = re.search(r'(\+?\d{7,})', contact_name)
        if match:
            extracted_num = match.group(1)
            # Check if it's the whole string
            if len(extracted_num) == len(contact_name.replace(" ", "")):
                # It's just a number
                contact_num = extracted_num
                # keep contact_name as is or same? 
                # If just number, name is number.
            else:
                # It is "Number Name" or "Name Number"
                contact_num = extracted_num
                contact_name = contact_name.replace(extracted_num, "").strip(" -.,")
                if not contact_name: 
                    contact_name = contact_num # Fallback if name becomes empty

    # Final cleanup: Separate Number from Name for Owner (fixing user request)
    if owner_name and owner_num == "":
        match = re.search(r'(\+?\d{7,})', owner_name)
        if match:
            extracted_num = match.group(1)
            if len(extracted_num) == len(owner_name.replace(" ", "")):
                owner_num = extracted_num
            else:
                owner_num = extracted_num
                owner_name = owner_name.replace(extracted_num, "").strip(" -.,")
                if not owner_name: owner_name = owner_num

    return owner_name, owner_num, contact_name, contact_num

    return owner_name, owner_num, contact_name, contact_num

def process_attachments(source_dir, output_dir):
    """
    Scans source_dir for files. 
    Copies them to output_dir/attachments.
    Returns mapping {filename: 'attachments/filename'}
    """
    mapping = {}
    att_dir = os.path.join(output_dir, "attachments")
    os.makedirs(att_dir, exist_ok=True)
    
    print(f"Scanning and copying attachments from {source_dir}...")
    
    exclude_prefix = os.path.abspath(output_dir)
    
    for root, dirs, files in os.walk(source_dir):
        if os.path.abspath(root).startswith(exclude_prefix):
            continue
            
        for file in files:
            # Skip the report itself or script
            if file in ["index.html", "generate_chat_report.py", "ChatReportGenerator.py", "chat_report_gui.py", "ChatReportGenerator_Pandas.py"]:
                continue
                
            src_path = os.path.join(root, file)
            # Flatten structure: copy all unique filenames to attachments/
            dest_path = os.path.join(att_dir, file)
            
            try:
                if not os.path.exists(dest_path):
                    shutil.copy2(src_path, dest_path)
                
                # Rel path for HTML
                mapping[file] = f"attachments/{file}"
            except Exception as e:
                pass
                
    return mapping

# ==========================================
# PARSERS (OpenPyXL Implementation)
# ==========================================

class BaseParser:
    def parse(self, filepath):
        raise NotImplementedError

class CellebriteParser(BaseParser):
    def parse(self, filepath):
        print("Using CellebriteParser (Light)...")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        
        sheet = None
        if 'Chat' in wb.sheetnames:
            sheet = wb['Chat']
        else:
            # Fallback to first
            sheet = wb.active
            
        chats = {}
        
        # Determine start row
        # Usually Cellebrite export has title then header. So row 1 is Title, Row 2 is Header, Row 3 is Data.
        # But indices in pandas were header=1 (so row 0 skip, row 1 header).
        # We will assume data starts from row 3 (index 3 in openpyxl) if checking pandas logic.
        # Let's start iterating from row 2 and see if it looks like data.
        
        # We'll skip first 2 rows.
        rows = list(sheet.iter_rows(values_only=True))
        
        # Heuristic: skip rows until we find one with an ID or looks like data
        # Pandas `header=1` means:
        # 0: "Report generated..."
        # 1: "ID", "Source", "Type"... (Header)
        # 2: 1, "WhatsApp", ... (Data)
        
        data_rows = rows[2:] # Slice off 0 and 1
        
        for idx, row in enumerate(data_rows):
            try:
                if not row or len(row) < 51: # Ensure row has minimal length
                    continue
                    
                # Index mapping (Pandas 0-based -> row[i])
                # 1: Chat ID
                chat_id = row[1]
                if chat_id is None: continue
                
                if chat_id not in chats:
                    chats[chat_id] = {
                        "id": chat_id,
                        "participants": str(row[9]) if row[9] is not None else "",
                        "owner": str(row[12]) if row[12] is not None else "",
                        "messages": []
                    }
                
                sender = str(row[21]) if row[21] is not None else ""
                owner_val = str(row[12]) if row[12] is not None else ""
                
                # Logic to determine is_sent
                normalized_sender = re.sub(r'[^0-9]', '', str(sender))
                normalized_owner = re.sub(r'[^0-9]', '', str(owner_val))
                is_sent = False
                if normalized_owner and normalized_owner in normalized_sender:
                    is_sent = True
                elif "owner" in sender.lower() or "proprietario" in sender.lower():
                    is_sent = True
                elif sender == owner_val:
                    is_sent = True
                
                body = clean_text(row[28])
                timestamp = str(row[37]) if row[37] is not None else ""
                
                att_name = row[45]
                att_info = str(att_name).strip() if att_name is not None and str(att_name).strip() != "" else None
                
                trans_raw = str(row[50]) if row[50] is not None else ""
                translation = ""
                if "Traduzione:" in trans_raw:
                    try:
                        translation = trans_raw.split("Traduzione:", 1)[1]
                        translation = clean_text(translation)
                    except: pass
                
                if body.strip().lower() in ["ok", "ok.", "ok!"] and len(translation) > 10:
                    translation = ""
                elif len(body) < 30 and len(translation) > (len(body) * 3 + 10):
                    translation = ""

                chats[chat_id]["messages"].append({
                    "sender": sender,
                    "is_sent": is_sent,
                    "body": body,
                    "time": timestamp,
                    "att": att_info,
                    "trans": translation
                })
            except Exception as e:
                # print(f"Row error: {e}")
                pass
        return chats

class WhatsAppParser(BaseParser):
    def parse(self, filepath, attachment_lookup=None):
        print(f"Using WhatsAppParser (Light)... Lookup size: {len(attachment_lookup) if attachment_lookup else 0}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        # Try multiple variations of the sheet name
        sheet_name = 'Instant Messages'
        if 'Messaggi istantanei' in wb.sheetnames:
            sheet_name = 'Messaggi istantanei'
        elif 'Instant Messages' in wb.sheetnames:
            sheet_name = 'Instant Messages'
        else:
            # Fallback: try to find any sheet with "Instant" or "Messaggi"
            found = False
            for name in wb.sheetnames:
                if "instant" in name.lower() or "messaggi" in name.lower():
                    sheet_name = name
                    found = True
                    break
            
            if not found:
                 raise KeyError("Sheet 'Messaggi istantanei' or 'Instant Messages' not found.")

        print(f"Reading sheet: {sheet_name}")
        sheet = wb[sheet_name]
        
        # Find headers
        rows = list(sheet.iter_rows(values_only=True))
        if not rows: 
            print("DEBUG: No rows found in sheet.")
            return {}
        
        if not rows: return {}
        
        # Search for header in first 10 rows
        header = None
        start_row_index = 0
        
        # Potential header keywords to look for
        keywords = ["Da", "A", "From", "To", "Corpo", "Body"]
        
        for i, row in enumerate(rows[:10]):
            row_vals = [str(c).lower() for c in row if c is not None]
            # Check if at least 2 keywords are in this row
            matches = sum(1 for k in keywords if k.lower() in row_vals)
            if matches >= 2:
                header = row
                start_row_index = i + 1
                print(f"Header found at row {i}: {header}")
                break
        
        if not header:
            # Fallback to row 0 if no header found (unlikely but safe)
            header = rows[0]
            start_row_index = 1
            print("Warning: Header not auto-detected, using first row.")

        # Map columns (handle both English and Italian)
        col_map = {name: i for i, name in enumerate(header) if name}
        
        chats = {}
        
        for row in rows[start_row_index:]:
            try:
                def get_col(names):
                    if isinstance(names, str): names = [names]
                    for n in names:
                        if n in col_map and col_map[n] < len(row):
                            return row[col_map[n]]
                    return None

                from_val = get_col(['From', 'Da'])
                to_val = get_col(['To', 'A'])
                
                if from_val is None and to_val is None: continue

                body_val = get_col(['Body', 'Corpo'])
                
                # Timestamp might be split or named differently
                time_val = get_col(['Timestamp-Time', 'Timestamp-Ora', 'Timestamp'])
                
                # Tag/Label for translation
                tag_val = get_col(['Tag', 'Etichetta', 'Note investigative', 'Messaggio Note investigative'])

                if from_val is None and to_val is None: continue

                direction_val = get_col(['Direction', 'Orientamento', 'Type', 'Tipo'])
                
                # Source info for attachments fallback
                source_info_val = get_col(['Informazioni sul file di origine', 'Source File Information'])

                # Participants logic
                sender_raw = str(from_val)
                receiver_raw = str(to_val)
                
                def clean_participant(p):
                    # Remove " - Inviato:..." or " - Sent:..." metadata
                    p = re.split(r'\s-\s(?:Inviato|Sent|Letti|Read|Delivered):', p, flags=re.IGNORECASE)[0]
                    # Remove (proprietario)/(owner)
                    p = re.sub(r'\((?:proprietario|owner|device owner)\)', '', p, flags=re.IGNORECASE)
                    # Remove _x000d_ and unicode garbage
                    p = p.replace('_x000d_', ' ').replace('\u200e', '').replace('\u202c', '')
                    # Replace newlines with space
                    p = p.replace('\n', ' ').replace('\r', ' ')
                    # Normalize whitespace
                    p = re.sub(r'\s+', ' ', p).strip()
                    # Remove trailing punctuation
                    p = p.strip('.,&;- ')
                    return p

                p1_clean = clean_participant(sender_raw)
                p2_clean = clean_participant(receiver_raw)
                
                # Filter out "Sconosciuto" if possible
                def is_valid_name(n):
                    if not n: return False
                    n_low = n.lower()
                    if "sconosciuto" in n_low or "unknown" in n_low: return False
                    if n_low == "&": return False
                    return True

                # Determine who is the "Owner" (Me)
                is_p1_owner = "(proprietario)" in sender_raw.lower() or "(owner)" in sender_raw.lower()
                is_p2_owner = "(proprietario)" in receiver_raw.lower() or "(owner)" in receiver_raw.lower()
                
                # Chat ID Generation
                parts = set(re.findall(r'(\d+@s\.whatsapp\.net)', sender_raw + " " + receiver_raw))
                if parts:
                    chat_id = "Chat " + "-".join(sorted([p.split('@')[0] for p in parts]))
                else:
                    # Use cleaned names for ID
                    parts_list = sorted([p for p in [p1_clean, p2_clean] if is_valid_name(p)])
                    if not parts_list:
                        # Fallback to whatever we have if both are sconosciuto
                        parts_list = sorted([p for p in [p1_clean, p2_clean] if p])
                        if not parts_list: parts_list = ["Unknown"]
                    
                    chat_id = f"Chat {' & '.join(parts_list)}"
                    chat_id = re.sub(r'[^\w\s\-\.]', '', chat_id)[:100]

                # Determine nice title for THIS row
                row_title = ""
                valid_names = [p for p in [p1_clean, p2_clean] if is_valid_name(p)]
                
                if is_p1_owner and p2_clean:
                    row_title = p2_clean
                elif is_p2_owner and p1_clean:
                    row_title = p1_clean
                elif len(valid_names) == 1:
                    row_title = valid_names[0]
                elif len(valid_names) >= 2:
                    row_title = " & ".join(valid_names)
                else:
                    # Both invalid/sconosciuto
                    row_title = "Chat Sconosciuta"

                if chat_id not in chats:
                    chats[chat_id] = {
                        "id": chat_id,
                        "participants": row_title, 
                        "owner": "Unknown",
                        "messages": []
                    }
                else:
                    # Update title if current one is bad and new one is better
                    curr_title = chats[chat_id]["participants"]
                    if ("sconosciuta" in curr_title.lower() or "unknown" in curr_title.lower()) and \
                       ("sconosciuta" not in row_title.lower() and "unknown" not in row_title.lower()):
                            chats[chat_id]["participants"] = row_title
                
                # Update Owner if found
                current_owner = chats[chat_id]["owner"]
                if current_owner == "Unknown":
                    if is_p1_owner:
                        chats[chat_id]["owner"] = p1_clean
                    elif is_p2_owner:
                        chats[chat_id]["owner"] = p2_clean

                # Sender Logic
                sender = p1_clean
                if "@s.whatsapp.net" in sender_raw:
                    m = re.search(r'(\d+@s\.whatsapp\.net)', sender_raw)
                    if m: sender = m.group(1).split('@')[0]
                
                # Direction Logic
                is_sent = False 
                
                # 1. Try explicit 'Direction' column
                if direction_val:
                    d_str = str(direction_val).lower()
                    if "uscita" in d_str or "outgoing" in d_str:
                        is_sent = True
                    elif "entrata" in d_str or "incoming" in d_str:
                        is_sent = False
                    # Else fall through
                
                # 2. Fallback to owner tag logic if direction ambiguous or not found
                if not is_sent: # Only check if not already confirmed sent
                     if is_p1_owner and p1_clean == sender: 
                        is_sent = True

                
                body = clean_text(body_val)
                ts_str = str(time_val) if time_val is not None else ""
                
                # Attachment Processing
                att_info = None
                source_info = str(source_info_val) if source_info_val else ""
                
                # Check for "part...mms" in source info (Signal/WhatsApp native dumps)
                if "part" in source_info and ".mms" in source_info:
                     # Attempt to extract part filename
                     try:
                         # Ex: .../app_parts/part123.mms/part123.mms.
                         # Logic: look for part<digits>.mms or similar
                         m = re.search(r'(part\d+\.mms)', source_info)
                         if m:
                             att_info = m.group(1) + "_" # Often has trailing underscore on disk
                         else:
                             att_info = "Attachment (unknown)"
                     except: pass
                
                # CROSS-REFERENCE LOOKUP
                if attachment_lookup and ts_str:
                     # Clean timestamp to match key
                     # Expect key to be: "20/05/2025 21:45:46" (roughly)
                     # Valid keys in map are strings.
                     # Remove (UTC...)
                     ts_clean = re.sub(r'\(UTC.*?\)', '', ts_str).strip()
                     
                     if ts_clean in attachment_lookup:
                         real_att = attachment_lookup[ts_clean]
                         if real_att:
                             att_info = real_att # OVERRIDE with real name
                
                # Translation extraction from Tag
                trans_raw = str(tag_val) if tag_val is not None else ""
                translation = ""
                
                # Filter out known non-translation tags
                ignore_tags = ["deleted", "read", "delivered", "sent", "cancellato", "letto", "consegnato", "inviato", "none", "nan"]
                
                if "Traduzione:" in trans_raw:
                    try:
                        translation = trans_raw.split("Traduzione:", 1)[1]
                        # Often followed by "Etichette:" or other fields
                        if "Etichette:" in translation:
                            translation = translation.split("Etichette:", 1)[0]
                        translation = clean_text(translation)
                    except: pass
                elif "Descrizione:" in trans_raw:
                    try:
                        translation = trans_raw.split("Descrizione:", 1)[1]
                        # Often followed by "Etichette:", "Creato:", "Modificato:"
                        if "Etichette:" in translation:
                            translation = translation.split("Etichette:", 1)[0]
                        if "Creato:" in translation:
                            translation = translation.split("Creato:", 1)[0]
                        translation = clean_text(translation)
                    except: pass
                elif trans_raw.strip() and trans_raw.lower() not in ignore_tags:
                    # Fallback: Assume the whole tag is a translation if it's not a status tag
                    # But exclude simple numbers or short codes
                    if len(trans_raw) > 2 and not trans_raw.isdigit():
                         translation = clean_text(trans_raw)

                chats[chat_id]["messages"].append({
                    "sender": sender,
                    "is_sent": is_sent,
                    "body": body,
                    "time": ts_str,
                    "att": att_info,
                    "trans": translation
                })
                
            except Exception as e:
                # print(f"Row error: {e}")
                pass
        
        return chats

# ==========================================
# RENDERERS
# ==========================================

class HTMLRenderer:
    def __init__(self, css, file_map, output_dir, style_mode="signal"):
        self.css = css
        self.file_map = file_map
        self.output_dir = output_dir
        self.style_mode = style_mode

    def render(self, chats, filename="index.html"):
        sidebar_html = ""
        chats_html = ""
        
        sorted_chat_ids = sorted(chats.keys())
        first = True
        
        for chat_id in sorted_chat_ids:
            data = chats[chat_id]
            msgs = data["messages"]

            # SORT MESSAGES CHRONOLOGICALLY
            def parse_ts(m):
                ts_str = m.get("time", "")
                if not ts_str: return datetime.min
                # Remove (UTC...)
                clean = re.sub(r'\(UTC.*?\)', '', ts_str).strip()
                try:
                    # Try common formats
                    # 20/05/2025 19:00:03
                    return datetime.strptime(clean, "%d/%m/%Y %H:%M:%S")
                except:
                    try:
                        return datetime.strptime(clean, "%d/%m/%Y %H:%M")
                    except:
                        try:
                            # Try just date? or different sep
                            return datetime.strptime(clean, "%Y-%m-%d %H:%M:%S")
                        except:
                            return datetime.min # Fail safe

            msgs.sort(key=parse_ts)

            # Parse Participants for Layout
            owner_name, owner_num, contact_name, contact_num = parse_participants_intelligent(data, chat_id)
            
            # Sidebar Item
            avatar_letter = contact_name[0].upper() if contact_name else "?"
            
            # Dynamic Owner Avatar
            owner_initial = owner_name[0].upper() if owner_name and owner_name != "Unknown" else "P"
            
            # Search Index Generation
            search_text = f"{contact_name} {contact_num} " + " ".join([f"{m['body']} {m['trans'] or ''}" for m in msgs])
            search_text = search_text.replace('"', '&quot;').lower()
            
            # SIDEBAR
            sidebar_html += f"""
            <div id="item-{chat_id}" class="chat-item" onclick="loadChat('{chat_id}')" data-search="{search_text}">
                <div class="sidebar-avatar">{avatar_letter}</div>
                <div style="flex:1;">
                    <div class="sidebar-name">{contact_name[:30]}</div>
                    <div class="sidebar-info">{len(msgs)} messaggi</div>
                </div>
            </div>
            """
            
            # Chat Area
            display_style = "flex" if first else "none"
            first = False
            
            msgs_html = ""
            current_date_str = None
            
            for idx, msg in enumerate(msgs):
                # Unique Message ID
                msg_id = f"msg-{chat_id}-{idx}"
                # Date Divider Logic
                try:
                    msg_date = msg['time'][:10] 
                    if msg_date != current_date_str and len(msg_date) >= 8 and ('/' in msg_date or '-' in msg_date):
                        current_date_str = msg_date
                        msgs_html += f'<div class="date-divider">{current_date_str}</div>'
                except: pass

                msg_class = "sent" if msg["is_sent"] else "received"
                
                sender_disp = msg["sender"]
                if msg["is_sent"]: 
                    sender_disp = owner_name if owner_name and owner_name != "Unknown" else "Tu"
                else:
                    # Received: prefer contact name over number
                    sender_disp = contact_name if contact_name else sender_disp

                sender_div = f'<div class="sender-name">{sender_disp}</div>'
                
                # Attachment
                att_html = ""
                if msg["att"]:
                    basename = os.path.basename(msg["att"])
                    if basename in self.file_map:
                        rel_path = self.file_map[basename]
                        _, ext_raw = os.path.splitext(basename)
                        ext = ext_raw.replace('.', '').lower().strip()
                        
                        img_exts = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg', 'ico', 'tif', 'tiff', 'heic', 'heif', 'mms', 'mms_']
                        vid_exts = ['mp4', 'webm', 'ogg', 'mov', 'qt', 'avi', 'mkv', 'm4v', '3gp']
                        aud_exts = ['mp3', 'wav', 'aac', 'm4a', 'amr', 'opus', 'oga', 'flac']

                        if ext in img_exts:
                            att_html = f'''
                            <div class="attachment">
                                <img src="{rel_path}" onclick="window.open(this.src)" title="{basename}" onerror="this.style.display='none'; this.parentNode.innerHTML='<a href={rel_path} target=_blank>📷 {basename} (Caricamento fallito)</a>'">
                            </div>
                            '''
                        elif ext in vid_exts:
                            att_html = f'''
                            <div class="attachment">
                                <video controls style="max-width:300px; max-height:300px; border-radius:10px;">
                                    <source src="{rel_path}">
                                    <a href="{rel_path}">📹 {basename}</a>
                                </video>
                            </div>
                            '''
                        elif ext in aud_exts:
                            att_html = f'''
                            <div class="attachment">
                                <audio controls style="max-width:300px;">
                                    <source src="{rel_path}">
                                    <a href="{rel_path}">🎵 {basename}</a>
                                </audio>
                            </div>
                            '''
                        else:
                            att_html = f'<div class="attachment"><a href="{rel_path}" target="_blank">📄 {basename}</a></div>'
                
                trans_html = ""
                if msg["trans"]:
                    trans_html = f'<div class="translation">{msg["trans"]}</div>'
                
                has_att_class = " has-attachment" if msg["att"] else ""
                msgs_html += f"""
                <div id="{msg_id}" class="message {msg_class}{has_att_class}">
                    {sender_div}
                    <div class="msg-content">
                        {msg['body']}
                        {att_html}
                    </div>
                    {trans_html}
                    <div class="timestamp">{msg['time']}</div>
                </div>
                """
                
            # Header
            owner_num_disp = owner_num if owner_num else "(Numero non presente)"
            contact_num_disp = contact_num if contact_num else "(Numero non presente)"

            header_html = f"""
            <div class="chat-info-header">
                <div class="participant-card">
                    <div class="participant-avatar owner">{owner_initial}</div>
                    <div class="participant-details">
                        <span class="participant-name">{owner_name}</span>
                        <span class="participant-number">{owner_num_disp}</span>
                    </div>
                </div>
                
                <div style="font-size:24px; color:#bbb">&#8596;</div>
                
                <div class="participant-card">
                    <div class="participant-avatar">{avatar_letter}</div>
                    <div class="participant-details">
                        <span class="participant-name">{contact_name}</span>
                        <span class="participant-number">{contact_num_disp}</span>
                    </div>
                </div>
            </div>
            """
            
            chats_html += f"""
            <div id="chat-{chat_id}" class="chat-messages" style="display:{display_style};">
                {header_html}
                {msgs_html}
            </div>
            """
            
        # Decide Sidebar Header based on Style
        if self.style_mode == "whatsapp":
            sidebar_header_div = """
            <div class="chat-header-bar">
                <div class="participant-avatar owner" style="width:40px;height:40px;font-size:16px;border:0;">T</div>
                <span style="font-weight:600; margin-left:15px; font-size:16px;">WhatsApp Report</span>
            </div>
            <div class="search-box">
                <input type="text" class="search-input" placeholder="Cerca messaggi..." onkeyup="performSearch(this.value)">
            </div>
            """
        else:
            # Default Signal/Generic
            sidebar_header_div = f'<div class="signal-sidebar-header">Chats ({len(chats)})</div>'
            sidebar_header_div += f'''
            <div class="search-box">
                <input type="text" class="search-input" placeholder="Cerca messaggi..." onkeyup="performSearch(this.value)">
            </div>
            '''
            
        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>{self.css}</style>
            <script>
                function loadChat(chatId, msgId) {{
                    // Hide all chat views
                    document.querySelectorAll('.chat-messages').forEach(e => e.style.display = 'none');
                    
                    // Show target chat
                    var chatView = document.getElementById('chat-' + chatId);
                    if(chatView) chatView.style.display = 'flex';
                    
                    // Update Active Sidebar Item
                    document.querySelectorAll('.chat-item').forEach(e => e.classList.remove('active'));
                    var item = document.getElementById('item-' + chatId);
                    if(item) item.classList.add('active');
                    
                    // If msgId provided, scroll and highlight
                    if(msgId) {{
                        var msgEl = document.getElementById(msgId);
                        if(msgEl) {{
                            msgEl.scrollIntoView({{behavior: 'smooth', block: 'center'}});
                            msgEl.classList.remove('highlight-msg'); // reset
                            void msgEl.offsetWidth; // trigger reflow
                            msgEl.classList.add('highlight-msg');
                        }}
                    }}
                }}
                
                function performSearch(query) {{
                    var chatList = document.getElementById('chat-list');
                    var resultsContainer = document.getElementById('search-results');
                    var resultsContent = document.getElementById('search-results-content');
                    
                    if (!query || query.length < 2) {{
                        chatList.style.display = 'block';
                        resultsContainer.style.display = 'none';
                        return;
                    }}
                    
                    chatList.style.display = 'none';
                    resultsContainer.style.display = 'block';
                    resultsContent.innerHTML = '';
                    
                    var q = query.toLowerCase();
                    var count = 0;
                    var maxResults = 100;
                    
                    // Search Logic: Iterate all messages in DOM
                    // Optimization: We could use the data-search on chats to filter first?
                    // But user wants specific messages.
                    
                    var allMessages = document.querySelectorAll('.message');
                    for (var i = 0; i < allMessages.length; i++) {{
                        if(count >= maxResults) break;
                        var msg = allMessages[i];
                        
                        // Check content
                        var text = msg.innerText; // innerText includes sender, time, body, translation
                        if(text.toLowerCase().includes(q)) {{
                            count++;
                            
                            // Context
                            var chatContainer = msg.closest('.chat-messages');
                            var chatId = chatContainer.id.replace('chat-', '');
                            var msgId = msg.id;
                            
                            // Get Chat Name from sidebar item
                            var chatItem = document.getElementById('item-' + chatId);
                            var chatName = chatItem ? chatItem.querySelector('.sidebar-name').innerText : 'Chat ' + chatId;
                            
                        // Highlight snippet
                        var lowerText = text.toLowerCase();
                        var idx = lowerText.indexOf(q);
                        var start = Math.max(0, idx - 40);
                        var end = Math.min(text.length, idx + 60);
                        var rawSnippet = text.substring(start, end);
                        if (start > 0) rawSnippet = "..." + rawSnippet;
                        if (end < text.length) rawSnippet = rawSnippet + "...";
                        
                        // Highlight term visually
                        // Escape regex special chars
                        var qEscaped = q.replace(/[.*+?^${{}}()|[\\]\\\\]/g, '\\\\$&');
                        var regex = new RegExp("(" + qEscaped + ")", "gi");
                        var styledSnippet = rawSnippet.replace(regex, '<span class="highlight-term">$1</span>');
                        
                        // Avatar styling
                        var avatarLetter = chatName.charAt(0).toUpperCase();
                        
                        var div = document.createElement('div');
                        div.className = 'search-result-item';
                        div.innerHTML = `
                            <div class="search-avatar">${{avatarLetter}}</div>
                            <div class="search-content">
                                <div class="search-result-sender">${{chatName}}</div>
                                <div class="search-result-preview">${{styledSnippet}}</div>
                            </div>
                        `;
                        
                        // Closure for click
                        (function(cid, mid) {{
                            div.onclick = function() {{ loadChat(cid, mid); }};
                        }})(chatId, msgId);
                        
                        resultsContent.appendChild(div);
                        }}
                    }}
                    
                    if (count === 0) {{
                        resultsContent.innerHTML = '<div style="padding:20px;text-align:center;color:#888">Nessun messaggio trovato.</div>';
                    }}
                }}
                
                function closeSearch() {{
                    document.querySelector('.search-input').value = '';
                    performSearch('');
                }}
            </script>
        </head>
        <body>
            <div class="container">
                <div class="sidebar">
                    {sidebar_header_div}
                    <div id="chat-list" style="flex:1; overflow-y:auto;">
                        {sidebar_html}
                    </div>
                    <div id="search-results" class="search-results-container">
                        <div class="search-header">
                            <span>RISULTATI RICERCA</span>
                            <span class="close-search" onclick="closeSearch()">CHIUDI</span>
                        </div>
                        <div id="search-results-content"></div>
                    </div>
                </div>
                <div class="main">
                    {chats_html}
                </div>
            </div>
        </body>
        </html>
        """
        
        path = os.path.join(self.output_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(full_html)
        print(f"Generated: {path}")

# ==========================================
# MAIN LOGIC
# ==========================================

def build_chat_attachment_map(filepath):
    """
    Reads 'Chat' sheet to build a map of Timestamp -> Attachment Filename.
    Used to resolve attachments in 'Instant Messages'.
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if 'Chat' not in wb.sheetnames: return {}
        sheet = wb['Chat']
        rows = list(sheet.iter_rows(values_only=True))
        
        # Simple/Safe header detection
        header = None
        start_row = 0
        for i, row in enumerate(rows[:5]):
            r_str = [str(x).lower() for x in row if x]
            # Key checks: 'timestamp: ora' (split), 'corpo', 'allegato'
            matches = 0
            if any("timestamp" in x for x in r_str): matches += 1
            if any("corpo" in x for x in r_str) or any("body" in x for x in r_str): matches += 1
            if matches >= 2:
                header = row
                start_row = i + 1
                break
        
        if not header: return {}
        
        col_map = {str(n).lower(): idx for idx, n in enumerate(header) if n}
        
        # Identify indices
        c_time_idx = -1
        # 'Timestamp: Ora' usually contains the full date/time string "20/05/2025 21:45:46(UTC+1)"
        for k, v in col_map.items():
            if "timestamp" in k and "ora" in k: c_time_idx = v; break
        
        c_att_idx = -1
        for k, v in col_map.items():
            if "allegato" in k and "#1" in k and "dettagli" not in k: c_att_idx = v; break
        
        if c_time_idx == -1 or c_att_idx == -1: return {}
        
        lookup = {}
        for row in rows[start_row:]:
            if len(row) <= max(c_time_idx, c_att_idx): continue
            
            ts_val = str(row[c_time_idx]) if row[c_time_idx] else ""
            att_val = str(row[c_att_idx]) if row[c_att_idx] else ""
            
            if ts_val and att_val:
                # Clean timestamp: remove newlines, (UTC...), whitespace
                ts_clean = re.sub(r'\(UTC.*?\)', '', ts_val).strip()
                lookup[ts_clean] = att_val
                
        return lookup
    except Exception as e:
        print(f"Error building attachment map: {e}")
        return {}

def run_generation(input_file, style="signal", source_type="chats", output_dir=None, format_type="auto"):
    input_path = os.path.abspath(input_file)
    base_dir = os.path.dirname(input_path)
    
    if output_dir:
        out_dir = output_dir
    else:
        out_dir = os.path.join(base_dir, "VisualReport")
        
    os.makedirs(out_dir, exist_ok=True)
    
    print(f"Input: {input_path}")
    print(f"Output: {out_dir}")
    print(f"Style: {style}")
    print(f"Source Mode: {source_type}")

    # 1. Determine Parser based on Source Type
    parser_impl = None
    att_lookup = None
    
    if source_type == "instant":
        print("Using Instant Messages Parser...")
        # Build attachment lookup from 'Chat' tab
        print("Building Attachment Lookup from 'Chat' tab...")
        att_lookup = build_chat_attachment_map(input_path)
        print(f"Attachment Lookup Built: {len(att_lookup)} entries.")
        
        parser_impl = WhatsAppParser() # Uses 'Instant Messages' tab logic
        chats = parser_impl.parse(input_path, attachment_lookup=att_lookup)
    else:
        print("Using Standard Chat Parser...")
        parser_impl = CellebriteParser() # Uses 'Chat' tab logic
        chats = parser_impl.parse(input_path)
    print(f"Parsed {len(chats)} chats.")
    
    # 3. Scan attachments
    file_map = process_attachments(base_dir, out_dir)
    
    # 4. Render
    css = CSS_SIGNAL if style == "signal" else CSS_WHATSAPP
    renderer = HTMLRenderer(css, file_map, out_dir, style_mode=style)
    renderer.render(chats)
    print("Generation Complete.")
    return out_dir

# ==========================================
# GUI CLASSES
# ==========================================

class RedirectText:
    def __init__(self, text_ctrl):
        self.output = text_ctrl

    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)

    def flush(self):
        pass

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Chat Report Generator (Lite)")
        self.root.geometry("600x500")
        
        # Styles
        self.root.configure(bg="#f0f0f0")
        font_label = ("Segoe UI", 10)
        
        # 1. File Selection
        frame_file = tk.Frame(root, bg="#f0f0f0", pady=10)
        frame_file.pack(fill=tk.X, padx=20)
        
        tk.Label(frame_file, text="File Excel:", bg="#f0f0f0", font=font_label).pack(anchor="w")
        
        self.entry_file = tk.Entry(frame_file, width=50)
        self.entry_file.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)
        
        btn_browse = tk.Button(frame_file, text="Sfoglia...", command=self.browse_file, font=font_label)
        btn_browse.pack(side=tk.RIGHT, padx=5)
        
        # 2. Options
        frame_opts = tk.Frame(root, bg="#f0f0f0", pady=10)
        frame_opts.pack(fill=tk.X, padx=20)
        
        # Style
        tk.Label(frame_opts, text="Stile Grafico:", bg="#f0f0f0", font=font_label).grid(row=0, column=0, padx=5, sticky="w")
        self.var_style = tk.StringVar(value="signal")
        tk.Radiobutton(frame_opts, text="Signal (Blue)", variable=self.var_style, value="signal", bg="#f0f0f0").grid(row=0, column=1)
        tk.Radiobutton(frame_opts, text="WhatsApp (Green)", variable=self.var_style, value="whatsapp", bg="#f0f0f0").grid(row=0, column=2)
        
        # Source
        tk.Label(frame_opts, text="Sorgente:", bg="#f0f0f0", font=font_label).grid(row=1, column=0, padx=5, sticky="w", pady=5)
        self.var_source = tk.StringVar(value="chats")
        tk.Radiobutton(frame_opts, text="Chat (Standard)", variable=self.var_source, value="chats", bg="#f0f0f0").grid(row=1, column=1, sticky="w")
        tk.Radiobutton(frame_opts, text="Messaggi Istantanei", variable=self.var_source, value="instant", bg="#f0f0f0").grid(row=1, column=2, sticky="w")
        
        # 3. Action
        frame_action = tk.Frame(root, bg="#f0f0f0", pady=20)
        frame_action.pack(fill=tk.X, padx=20)
        
        btn_run = tk.Button(frame_action, text="GENERA REPORT", command=self.start_generation, 
                            bg="#2c6bed", fg="white", font=("Segoe UI", 12, "bold"), height=2)
        btn_run.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_open = tk.Button(frame_action, text="APRI CARTELLA OUTPUT", command=self.open_output, 
                                  state=tk.DISABLED, font=("Segoe UI", 10))
        self.btn_open.pack(fill=tk.X)
        
        # 4. Log Output
        tk.Label(root, text="Log:", bg="#f0f0f0", font=font_label).pack(anchor="w", padx=20)
        self.text_log = scrolledtext.ScrolledText(root, height=15)
        self.text_log.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        # Redirect stdout
        sys.stdout = RedirectText(self.text_log)
        sys.stderr = RedirectText(self.text_log)
        
        self.last_output_dir = None

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, filename)

    def start_generation(self):
        input_path = self.entry_file.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showwarning("Attenzione", "Seleziona un file Excel valido.")
            return
            
        style = self.var_style.get()
        source = self.var_source.get()
        
        # Disable button
        self.btn_open.config(state=tk.DISABLED)
        
        t = threading.Thread(target=self.run_process, args=(input_path, style, source))
        t.start()
        
    def run_process(self, input_path, style, source):
        print("--- Inizio Elaborazione ---")
        try:
            # Determine output automatically (same dir)
            out_dir = run_generation(input_path, style=style, source_type=source, output_dir=None)
            self.last_output_dir = out_dir
            
            self.root.after(0, lambda: self.finish_success())
        except Exception as e:
            error_msg = str(e)
            print(f"ERRORE: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror("Errore", f"Si è verificato un errore:\n{error_msg}"))
        finally:
            print("--- Fine Elaborazione ---")

    def finish_success(self):
        messagebox.showinfo("Successo", "Report generato con successo!")
        self.btn_open.config(state=tk.NORMAL)
        
    def open_output(self):
        if self.last_output_dir and os.path.exists(self.last_output_dir):
            os.startfile(self.last_output_dir)
        else:
            messagebox.showwarning("Errore", "Cartella non trovata.")

# ==========================================
# ENTRY POINT
# ==========================================

def main():
    # If args passed, run CLI mode
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description="Convert Excel Chat export to HTML")
        parser.add_argument("input_file", help="Path to Excel file")
        parser.add_argument("--style", choices=["signal", "whatsapp"], default="signal", help="Output visual style")
        parser.add_argument("--output", help="Output directory")
        parser.add_argument("--format", choices=["auto", "cellebrite", "whatsapp"], default="auto", help="Input format")
        
        args = parser.parse_args()
        run_generation(args.input_file, args.style, args.output, args.format)
    else:
        # GUI Mode
        root = tk.Tk()
        app = App(root)
        root.mainloop()

if __name__ == "__main__":
    main()
