# %% Cell 1
!pip install git+https://github.com/openai/whisper.git
!sudo apt update && sudo apt install ffmpeg

# %% Cell 2
!whisper "20241128_163155-Michael-Serva.mp3" --model medium --language Spanish

# %% Cell 3
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)
display(df)

# %% Cell 4
!whisper "20241128_160952-Michael-Serva.mp3" --model medium --language Spanish

# %% Cell 5
srt_file_1 = '/content/20241128_163155-Michael-Serva.srt'
srt_file_2 = '/content/20241128_160952-Michael-Serva.srt'
output_file = 'combined_transcript.srt'

with open(srt_file_1, 'r') as f1:
    content1 = f1.read()

with open(srt_file_2, 'r') as f2:
    content2 = f2.read()

with open(output_file, 'w') as outfile:
    outfile.write(content1)
    outfile.write('\n\n') # Add a separator between the two transcripts, if needed
    outfile.write(content2)

print(f"Combined SRT file saved to {output_file}")

# %% Cell 6
import pandas as pd, os, re, numpy as np

path="/content/combined_transcript.srt"
with open(path,'r',encoding='utf-8',errors='replace') as f:
    raw_lines=f.read().splitlines()

ts_re=re.compile(r'^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}')
num_re=re.compile(r'^\d+$')

utterances=[]
for line in raw_lines:
    if line=="":
        continue
    if num_re.match(line.strip()):
        continue
    if ts_re.match(line.strip()):
        continue
    utterances.append(line)

punct_re=re.compile(r'[^\wáéíóúüñÁÉÍÓÚÜÑ]+', re.UNICODE)
def tokenize(u): return [t for t in re.split(r'\s+', u.strip()) if t]
def norm_word(w):
    w=w.lower()
    w=punct_re.sub('', w)
    return w

ack_tokens=set([
    "ajá","aja","claro","ya","ok","okay","okei","sí","si","correcto","exacto","perfecto","bien","listo",
    "deacuerdo","de","acuerdo","mmm","mm","mhm","vale","dale","asi","así","seguro","totalmente","entendido","okey"
])
def ack_only(u):
    toks=[norm_word(t) for t in tokenize(u)]
    toks=[t for t in toks if t]
    if not toks: return False
    if toks==["bueno"]: return True
    return all(t in ack_tokens for t in toks)

def is_question(u): return ("?" in u) or ("¿" in u)
digit_re=re.compile(r'\d')
money_re=re.compile(r'(\$|s/|soles|dólar|dolar|usd|euros?)', re.I)
unit_re=re.compile(r'(tonelad|mill[oó]n|m3|kg|kwh|horas?|d[ií]as?|semanas?|meses?|años?|%|por\s+tonelada)', re.I)
def has_numbers(u): return bool(digit_re.search(u)) or bool(money_re.search(u))
def is_long(u): return len(tokenize(u))>=10 or len(u)>=70
def interviewer_cue(u):
    lu=u.strip().lower()
    return lu.startswith(("entonces","ya","ok","okay","a ver","veamos","podemos","pasemos","revisemos","quisiera","mira","perfecto"))

connectors=re.compile(r'^(que|porque|por que|y|adem[aá]s|incluso|luego|pero|bueno|o sea|más bien|también)\b', re.I)

def tag_question_informant(u):
    lu=u.strip().lower()
    if lu.startswith("¿"):
        return False
    return lu.endswith(("¿no?","¿cierto?","¿verdad?")) and (has_numbers(u) or unit_re.search(u))

roles=[]
prev_is_question=False
prev_speaker=None
for u in utterances:
    q=is_question(u)
    ack=ack_only(u)
    nums=has_numbers(u)
    long=is_long(u)
    cue=interviewer_cue(u)
    wc=len(tokenize(u))
    if ack:
        typ="I1" if (prev_is_question and prev_speaker=="INT") else "INT"
    else:
        if q and not tag_question_informant(u):
            typ="INT"
        else:
            if prev_is_question and prev_speaker=="INT":
                typ="I1"
            elif prev_speaker=="I1" and connectors.match(u.strip()) and wc<=12 and not q:
                typ="I1"
            elif long or nums or unit_re.search(u):
                typ="I1"
            elif cue and wc<=12:
                typ="INT"
            else:
                typ="I1"
    roles.append(typ)
    prev_is_question = (q and not tag_question_informant(u))
    prev_speaker = typ

final_roles=[]
for idx,(u,typ) in enumerate(zip(utterances,roles)):
    if typ=="I1":
        final_roles.append("I1")
    else:
        wc=len(tokenize(u))
        q=is_question(u)
        ack=ack_only(u)
        norm=" ".join([norm_word(t) for t in tokenize(u) if norm_word(t)])
        if ack or wc<=2:
            final_roles.append("E2")
        elif (not q) and wc<=4 and norm in {"ya","ok","okay","claro","perfecto","bien","listo","dale","correcto","exacto"}:
            final_roles.append("E2")
        elif q and wc<=4 and idx>0 and final_roles[idx-1]=="I1":
            final_roles.append("E2")
        else:
            final_roles.append("E1")

n=len(utterances)
width=max(3, len(str(n)))
line_ids=[f"L{str(i+1).zfill(width)}" for i in range(n)]
df=pd.DataFrame({"LÍNEA_ID": line_ids, "ROL": final_roles, "TEXTO": utterances})
out_path="/content/roles_clasificados.tsv"
df.to_csv(out_path, sep="\t", index=False, encoding="utf-8")
df.head(25), out_path, n

# %% Cell 7
import pandas as pd, os, re, numpy as np

path="/content/combined_transcript.srt"
with open(path,'r',encoding='utf-8',errors='replace') as f:
    raw_lines=f.read().splitlines()

ts_re=re.compile(r'^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}')
num_re=re.compile(r'^\d+$')
utterances=[]
for line in raw_lines:
    if line=="": continue
    if num_re.match(line.strip()): continue
    if ts_re.match(line.strip()): continue
    utterances.append(line)

punct_re = re.compile(r'[^\wáéíóúüñÁÉÍÓÚÜÑ]+', re.UNICODE)
def tokenize(u): return [t for t in re.split(r'\s+', u.strip()) if t]
def norm_word(w):
    w = w.lower()
    w = punct_re.sub('', w)
    return w
ack_tokens = set(["ajá","aja","claro","ya","ok","okay","okei","sí","si","correcto","exacto","perfecto","bien","listo","deacuerdo","de","acuerdo","mmm","mm","mhm","vale","dale","asi","así","seguro","totalmente","entendido","okey"])
def ack_only(u):
    toks=[norm_word(t) for t in tokenize(u)]
    toks=[t for t in toks if t]
    if not toks: return False
    if toks == ["bueno"]: return True
    return all(t in ack_tokens for t in toks)
def is_question(u): return ("?" in u) or ("¿" in u)
digit_re = re.compile(r'\d')
money_re = re.compile(r'(\$|s/|soles|dólar|dolar|usd|euros?)', re.I)
unit_re = re.compile(r'(tonelad|mill[oó]n|m3|kg|kwh|horas?|d[ií]as?|semanas?|meses?|años?|%|por\s+tonelada)', re.I)
def has_numbers(u): return bool(digit_re.search(u)) or bool(money_re.search(u))
def is_long(u): return len(tokenize(u)) >= 10 or len(u) >= 70
def interviewer_cue(u):
    lu=u.strip().lower()
    return lu.startswith(("entonces","ya","ok","okay","a ver","veamos","podemos","pasemos","revisemos","quisiera","mira","perfecto"))
connectors=re.compile(r'^(que|porque|por que|y|adem[aá]s|incluso|luego|pero|bueno|o sea|más bien|también)\b', re.I)
def tag_question_informant(u):
    lu=u.strip().lower()
    if lu.startswith("¿"): return False
    return lu.endswith(("¿no?","¿cierto?","¿verdad?")) and (has_numbers(u) or unit_re.search(u))
roles=[]
prev_is_question=False
prev_speaker=None
for u in utterances:
    q=is_question(u); ack=ack_only(u); nums=has_numbers(u); long=is_long(u); cue=interviewer_cue(u); wc=len(tokenize(u))
    if ack: typ="I1" if (prev_is_question and prev_speaker=="INT") else "INT"
    elif q and not tag_question_informant(u): typ="INT"
    elif prev_is_question and prev_speaker=="INT": typ="I1"
    elif prev_speaker=="I1" and connectors.match(u.strip()) and wc<=12 and not q: typ="I1"
    elif long or nums or unit_re.search(u): typ="I1"
    elif cue and wc<=12: typ="INT"
    else: typ="I1"
    roles.append(typ); prev_is_question=(q and not tag_question_informant(u)); prev_speaker=typ
final_roles=[]
for idx,(u,typ) in enumerate(zip(utterances,roles)):
    if typ=="I1": final_roles.append("I1"); continue
    wc=len(tokenize(u)); q=is_question(u); ack=ack_only(u); norm=" ".join([norm_word(t) for t in tokenize(u) if norm_word(t)])
    if ack or wc<=2: final_roles.append("E2")
    elif (not q) and wc<=4 and norm in {"ya","ok","okay","claro","perfecto","bien","listo","dale","correcto","exacto"}: final_roles.append("E2")
    elif q and wc<=4 and idx>0 and final_roles[idx-1]=="I1": final_roles.append("E2")
    else: final_roles.append("E1")
n=len(utterances); width=max(3,len(str(n))); line_ids=[f"L{str(i+1).zfill(width)}" for i in range(n)]
df=pd.DataFrame({"LÍNEA_ID":line_ids,"ROL":final_roles,"TEXTO":utterances})
out_path="/content/roles_clasificados.tsv"
df.to_csv(out_path,sep="\t",index=False,encoding="utf-8")
print("TSV generado en:", out_path)
print(df.head(10))

# %% Cell 8
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)
display(df)
