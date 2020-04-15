from docx import Document
import re
import json
import sys
import msoffcrypto
import os
import csv

def decrypt(files_folder, password = 'p65Lk!', save_folder = 'data/Cha_Lab_Transcripts_Decrypted/'):
    files = [f for f in os.listdir(files_folder) if '.docx' in f]
    for f in files:
        file_name = files_folder+f
        file = msoffcrypto.OfficeFile(open(file_name, "rb"))
        file.load_key(password=password)
        save_file_name = save_folder+f
        file.decrypt(open(save_file_name, "wb"))

class segment(object):
    def __init__(self, data=None, text=None, color=None, tag=None, pos=None):
        if data: self.load_data(data)
        elif text is not None and color is not None and tag is not None and pos is not None:
            self.text = text
            self.tag = tag.upper()
            self.check_input(color, self.tag)
            self.primary = self.extract_primary(color, self.tag)
            self.secondary = self.extract_secondary(color, self.tag)
            self.pos = pos
        else:
            print(text, color, tag, pos)
            raise Exception('Segment Error. Invalid input.', )

    def load_data(self, data):
        self.text = data['text']
        self.tag = data['tag'].upper()
        self.primary = data['primary']
        self.secondary = data['secondary']
        self.pos = data['pos']
        
    def __repr__(self): return ' {} ({})({}) '.format(self.text, self.primary, self.secondary)

    def __str__(self): return self.text

    def __len__(self): return len(self.text.split(' '))

    def check_input(self, color, tag):
        if color not in ['green', 'red', 'gray']: 
            raise Exception('Error. Color unrecognized. Color: {}'.format(color))
        if tag not in ['R', 'NR', 'E', 'PL', 'T', 'ET', 'PE', '', 'NA']:
            raise Exception('Segment Error. Tag unrecognized. Tag: {}'.format(tag))
        
    def extract_primary(self, color, tag):
        if color=='green': return 'internal'
        elif color=='red': return 'external'
        elif color=='gray' and tag in ['R', 'NR', '', 'NA']: return 'other'
        else: raise Exception('Segment Primary Error. Wrong color and tag. Color: {} Tag: {}'.format(color, tag))
            
    def extract_secondary(self, color, tag):
        if color in ['red', 'gray']: return ''
        else:
            if tag=='E': return 'event'
            elif tag=='PL': return 'place'
            elif tag=='T': return 'time'
            elif tag=='ET': return 'emotion'
            elif tag=='PE': return 'perceptual'
            elif tag in ['R', 'NR', 'NA', '']: return ''
            else: raise Exception('Segment Secondary Error. Tag unrecognized. Tag: {}'.format(tag))
                
    def save_data(self):
        return {'text':self.text, 'tag':self.tag, 'primary': self.primary, 'secondary': self.secondary, 'pos': self.pos}


class document(object):
    def __init__(self, data, docID=None, eventType=None):
        self.segmentsList = []
        self.WC, self.SC = 0, 0
        self.docID, self.eventType = docID, eventType
        self.primaries_count = {'internal':0, 'external':0, 'other':0}
        self.secondaries_count = {'event':0, 'place':0, 'time':0, 'emotion':0, 'perceptual':0}
        self.primaries = {'internal': [], 'external': [], 'other': []}
        self.secondaries = {'event': [], 'place': [], 'time': [], 'emotion': [], 'perceptual': []}
        self.logistics = None
        if type(data) == str: self.parse_file(data)
        elif type(data) == dict: self.parse_data(data)
        else: raise Exception('Document Error. Invalid Input.')
        
    def parse_file(self, file_path):
        if '.json' in file_path:
            self.parse_data(json.load(open(file_path)))
        elif '.docx' in file_path:
            if not self.eventType: self.eventType = file_path.split('/')[-1].split('_')[2].strip()
            if self.eventType not in ['pos', 'neg']: raise Exception('Document Error. Wrong Event Type.')
            if not self.docID: self.docID = file_path.split('/')[-1].split('_')[0].strip() + '_' + self.eventType
            d = Document(file_path)
            for para in d.paragraphs: self.add_segments(para)
        else: raise Exception('Document Error. Invalid load file format. File: {}'.format(file_path))
            
    def parse_data(self, data):
        if 'docID' not in data: raise Exception('Document Error. Cannot find docID.')
        if 'eventType' not in data: raise Exception('Document Error. Cannot find eventType.')
        self.docID,self.eventType = data['docID'], data['eventType']
        for d in data['data']:
            seg = segment(data=d)
            self.segmentsList.append(seg)
            self.calculate_seg_logistics(seg)
        self.logistics = self.calculate_logistics()

    def get_short_rep(self):
        return ' '.join([str(seg) for seg in self.segmentsList])[:80] + ' ...'
        
    def __repr__(self):
        return 'Type: DOCUMENT   ID: {}\n{}'.format(self.docID, self.get_short_rep())
    
    def __str__(self): return ' '.join([str(seg) for seg in self.segmentsList])

    def __len__(self): return self.WC
    
    def calculate_logistics(self):
        return {
            'WC': self.WC, 'SC': self.SC,
            'primaries': self.primaries_count, 'secondaries': self.secondaries_count }
    
    def print_logistics(self):
        print('''DOCUMENT\nWord Count: {}   Segment Count: {}\n
                    Primaries: {}\n Secondaries: {}\n'''.format(self.WC, self.SC,
                   '\t'.join(['{}:{}'.format(k,v) for k,v in self.primaries_count.items()]),
                   '\t'.join(['{}:{}'.format(k,v) for k,v in self.secondaries_count.items()])) )
    
    def extract_tag(self, text):
        spl, r = text, ''
        if text:
            reg = re.findall(r'\[[A-Za-z]*\]', text)
            if reg:
                r = reg[-1]
                spl = text.split(r)[0].strip()
        return spl.strip(), r.strip()
    
    def extract_text(self, paragraph):
        para_text = paragraph.text
        ans = {}
        for text in para_text.split('│'):
            spl,r = self.extract_tag(text.strip())
            if spl: ans[spl] = r
        return ans
        
    def extract_highlight(self, paragraph):
        color_mapping = {16: 'gray', 4: 'green', 6: 'red', 15: 'gray'}
        skip_color=[]
        curr_color, curr_text, curr_tag = '', '', ''
        details = [[]]
        for run in paragraph.runs:
            high, text = run.font.highlight_color, run.text
            if 'R:' in text: continue
            if high == None:
                if curr_text.strip():
                    spl, tag = self.extract_tag(curr_tag.replace('│', '').strip())
                    tag = tag.replace('[','').replace(']','')
                    details[-1].append(tag)
                    if curr_color not in skip_color:
                        details.append([curr_text.strip(), color_mapping[curr_color]])
                    curr_tag = ''
                curr_tag += text
                curr_color, curr_text = '', ''
            else:
                if high == curr_color: curr_text += text
                else:
                    if curr_text.strip():
                        spl, tag = self.extract_tag(curr_tag.replace('│', '').strip())
                        tag = tag.replace('[','').replace(']','')
                        details[-1].append(tag)
                        if curr_color not in skip_color:
                            details.append([curr_text.strip(), color_mapping[curr_color]])
                        curr_tag = ''
                    curr_color, curr_text = high, text
        spl, tag = self.extract_tag(curr_tag.replace('│', '').strip())
        tag = tag.replace('[','').replace(']','')
        details[-1].append(tag)
        return details[1:]
    
    def add_segments(self, para, method='highlight'):
        if method == 'highlight':
            raw_seg = self.extract_highlight(para)
            for entry in raw_seg:
                text, color, tag = entry[:]
                seg = segment(text=text, color=color, tag=tag, pos=len(self.segmentsList))
                self.segmentsList.append(seg)
                self.calculate_seg_logistics(seg)
            self.logistics = self.calculate_logistics()
        
    def calculate_seg_logistics(self, seg):
        self.WC +=len(seg)
        self.SC +=1
        self.primaries_count[seg.primary] += 1
        self.primaries[seg.primary].append(seg)
        if seg.secondary:
            self.secondaries_count[seg.secondary] += 1
            self.secondaries[seg.secondary].append(seg)
            
    def save(self, save_file=None):
        if not save_file: save_file = str(self.docID)+'.json'
        if '.json' not in save_file: raise Exception('Document Error. Wrong save file format.')
        segs_json = self.save_data()
        json.dump(segs_json, open(save_file, 'w'), indent=2)
    
    def save_data(self):
        return {'docID':self.docID, 'eventType': self.eventType, 'data':[d.save_data() for d in self.segmentsList]}

class participant(object):
    def __init__(self, files, SI_hx=None, SI_3mo=None, SI_6mo=None,  partID=None):
        self.partID, self.SI_hx, self.SI_3mo, self.SI_6mo = partID,SI_hx, SI_3mo, SI_6mo
        self.docs, self.docsList = {}, []
        self.WC, self.SC, self.DC = 0, 0, 0
        self.primaries_count = {'internal':0, 'external':0, 'other':0}
        self.secondaries_count = {'event':0, 'place':0, 'time':0, 'emotion':0, 'perceptual':0}
        self.primaries = {'internal': [], 'external': [], 'other': []}
        self.secondaries = {'event': [], 'place': [], 'time': [], 'emotion': [], 'perceptual': []}
        self.logistics = None
        if type(files) == list: self.parse_file(files)
        elif type(files) == str: self.parse_json(files)
        elif type(files)==dict: self.parse_data(files)
        else: raise Exception('Participant Error. Invalid input.')
            
    def __repr__(self):
        st = '\n'.join([doc.get_short_rep() for doc in self.docs.values()])
        return 'Type: PARTICIPANT   ID: {}   SI_hx: {}   SI_3mo: {}   SI_6mo: {}\n{}'\
            .format(self.partID, self.SI_hx, self.SI_3mo, self.SI_6mo, st)
    
    def __str__(self): return '\n'.join([doc.get_short_rep() for doc in self.docs.values()])

    def __len__(self): return self.DC
            
    def check_files_format(self, files, form):
        for file in files:
            if form not in file: return False
        return True
    
    def parse_file(self, files):
        if self.check_files_format(files, '.docx') or self.check_files_format(files, '.json'):
            if self.SI_hx is None: raise Exception('Participant Error. No SI_hx.')
            if self.SI_3mo is None: raise Exception('Participant Error. No SI_3mo.')
            if self.SI_6mo is None: raise Exception('Participant Error. No SI_6mo.')
            for file in files:
                if not self.partID: self.partID = file.split('/')[-1].split('_')[0].strip()
                else:
                    if self.partID != file.split('/')[-1].split('_')[0].strip(): raise Exception('Participant Error. PartID Error')
                doc_temp = document(file)
                self.docs[doc_temp.docID] = doc_temp
                self.docsList.append(doc_temp)
                self.calculate_doc_logistics(doc_temp)
            self.logistics = self.calculate_logistics()
        else: raise Exception('Participant Error. Invalid input files')


    def parse_json(self, files):
        if '.json' not in files: raise Exception('Participant Error. Invalid input files')
        data = json.load(open(files))
        self.parse_data(data)

    def parse_data(self, data):
        if 'partID' not in data: raise Exception('Participant Error. No partID.')
        if 'SI_hx' not in data: raise Exception('Participant Error. No SI_hx.')
        if 'SI_3mo' not in data: raise Exception('Participant Error. No SI_3mo.')
        if 'SI_6mo' not in data: raise Exception('Participant Error. No SI_6mo.')
        self.partID, self.SI_hx, self.SI_3mo, self.SI_6mo = data['partID'], data['SI_hx'], data['SI_3mo'], data['SI_6mo']
        for k,v in data['docs'].items():
            doc = document(v, k)
            self.docs[k] = document(v, k)
            self.docsList.append(doc)
            self.calculate_doc_logistics(doc)
        self.logistics = self.calculate_logistics()
            
    def calculate_doc_logistics(self, doc):
        self.DC +=1
        self.SC += doc.SC
        self.WC += doc.WC
        for p in self.primaries:
            self.primaries_count[p]+=doc.primaries_count[p]
            self.primaries[p].extend(doc.primaries[p])
        for s in self.secondaries:
            self.secondaries_count[s]+=doc.secondaries_count[s]
            self.secondaries[s].append(doc.secondaries[s])
            
    def calculate_logistics(self):
        return {
            'WC': self.WC, 'SC': self.SC, 'DC': self.DC,
            'primaries': self.primaries_count, 'secondaries': self.secondaries_count}
    
    def print_logistics(self):
        print('PARTICIPANT ID: {}   SI_hx: {}   SI_3mo: {}   SI_6mo: {}\nWord Count: {}\t   Segment Count: {}   Document Count: {}\n'
              'Primaries: {}\n Secondaries: {}\n'.format(self.partID, self.SI_hx, self.SI_3mo, self.SI_6mo, self.WC, self.SC, self.DC,\
               '\t'.join(['{}:{}'.format(k,v) for k,v in self.primaries_count.items()]), \
               '\t'.join(['{}:{}'.format(k,v) for k,v in self.secondaries_count.items()])) )
            
    def save(self, save_file=None):
        if not save_file: save_file = str(self.partID)+'.json'
        if '.json' not in save_file: raise Exception('Participant Error. Wrong save file format.')
        docs_json = self.save_data()
        json.dump(docs_json, open(save_file, 'w'), indent=2)
        
    def save_data(self):
        return {'partID': self.partID, 'SI_hx': self.SI_hx, 'SI_3mo': self.SI_3mo, 'SI_6mo': self.SI_6mo,
                'docs': {k:d.save_data() for k,d in self.docs.items()}}


# if __name__== '__main__':
#     args = sys.argv
#     if len(args)==1:
#         print('Loading Participant Data (placeholder values: SI=True, ID=123)\n------')
#         part = participant(['sample/event3.docx', 'sample/event1.docx'], SI=True, partID=123)
#         print(part)
#     else:
#         if args[1].strip() not in ['document', 'participant']: raise Exception('Invalid arguments')
#         arg = args[1].strip()
#         param = args[2:]
#         if arg=='document':
#             print('Loading Document {}\n------'.format(param[0]))
#             doc = document(param[0])
#             print(doc)
#         else:
#             print('Loading Participant Data (placeholder values: SI=True, ID=123)\n------')
#             part = participant(param, SI=True, partID=123)
#             print(part)

# decrypt('data/Cha_Lab_Transcripts/')

def save_participants_json(root_folder='data', participants_file_folder='Cha_Lab_Transcripts_Decrypted',
                      SI_file_path='Cha_Lab_ML Project_Subject_Variables_02 April 2020.csv', save_file='participants.json'):
    def convert_SI(inp):
        if int(inp) == 999: return -1
        else: return int(inp)

    participants = []
    folder_path = os.path.join(root_folder, participants_file_folder)
    SI_file_fullPath = os.path.join(root_folder, SI_file_path)
    save_file_path = os.path.join(root_folder, save_file)

    files = [os.path.join(folder_path,f) for f in os.listdir(folder_path) if '.docx' in f]
    files.sort(key=lambda x:(int(x.split('/')[-1].split('_')[0][1:]), x.split('/')[-1].split('_')[2]))
    SI_file = csv.reader(open(SI_file_fullPath))
    SI_file_data = [row for row in SI_file]
    count = 1
    for i in range(0,len(files),2):
        subject, SI_hx,SI_3mo,SI_6mo = SI_file_data[count]
        SI_hx, SI_3mo, SI_6mo = convert_SI(SI_hx), convert_SI(SI_6mo), convert_SI(SI_3mo)
        participants.append(participant(files[i:i+2], SI_hx,SI_3mo,SI_6mo).save_data())
        count+=1

    json.dump(participants, open(save_file_path,'w'), indent=2)

def load_participants_json(json_file='data/participants.json'):
    file = json.load(open(json_file))
    participants = []
    for f in file: participants.append(participant(f))
    return participants
