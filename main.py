import os,subprocess
import re
import argparse
import pandas as pd
from tqdm import tqdm
from glob import glob
from shutil import copy, move
from pathlib import Path

parser = argparse.ArgumentParser()
parser.add_argument('-p','--pck',type=str,help='Game Path for voice PCKs',default="C:/Program Files/HoYoPlay/games/Honkai Star Rail/Star Rail/Games")
parser.add_argument("-l", "--language", type=str, help="Language of the voice pack", required=True)
parser.add_argument('-w','--wav',type=str,help='Path for final wavs',default="./Data/wav")
parser.add_argument("-m", "--mode", type=str, default="mv", help="Move or copy files into wav. Defaults to move")
parser.add_argument('-dst','--destination', type=str, help='Final output location', default="./Data/FinalOutput")


args = parser.parse_args()


pck_path = args.pck + "/StarRail_Data/"
wem_path = r"./Data/raw"
wav_path = args.wav
language = args.language
final_destination = args.destination # smash reference lol

if not os.path.exists(wem_path):
    Path(wem_path).mkdir(parents=True)
if not os.path.exists(wav_path):
    Path(wav_path).mkdir(parents=True)

# part 1
def unpack(pck, wem):
    pck_file = glob(f"{pck}/**/*.pck", recursive=True)
    for pcks in tqdm(pck_file, desc="正在解包音频..."):
        subprocess.run(f"./Tools/quickbms.exe -q -k ./Tools/wwise_pck_extractor.bms \"{pcks}\" \"{wem}\"",
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def to_wav(wem, wav):
    wem_file = glob(f"{wem}/**/*.wem", recursive=True)
    for wems in tqdm(wem_file, desc="正在转码音频，请耐心等待完成..."):
        wav_name = os.path.basename(wems).replace(".wem", ".wav")
        subprocess.run(f"./Tools/vgmstream-cli.exe -o \"{wav}/{wav_name}\" \"{wems}\"", stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
    tqdm.write(f"Sucessfully unpacked {wem_file} to {wav}")


unpack(pck_path, wem_path)
to_wav(wem_path, wav_path)

#part 2
def crean_text(text):
    html_tag = re.compile(r'<.*?>')
    text = re.sub(html_tag, '', text)
    text = text.replace('\n', ' ')
    return text


def is_in(full_path, regx):
    if re.findall(regx, full_path):
        return True
    else:
        return False


def get_support_lang(lang):
    indexs = glob('./Indexs/*.xlsx')
    path = ['中文 - Chinese', '英语 - English', '日语 - Japanese', '韩语 - Korean']
    support_langs = []
    for langs in indexs:
        lang_code = Path(langs).name.replace(".xlsx", "")
        support_langs.append(lang_code)
    if lang in support_langs:
        langcodes = support_langs
        i = langcodes.index(lang)
        dest_path = path[i]
    else:
        print("Invalid Language! Defaulting to EN")
        dest_path = "英语 - English"
    return lang, dest_path


def sorting_voice(src, dst, mode, lang):
    langcode, dest_path = get_support_lang(lang)
    df = pd.read_excel(f"./Indexs/{langcode}.xlsx")
    for i, row in tqdm(df.iterrows(), desc="正在整理数据集...", total=len(df)):
        character = row['角色']
        voice_hash = row['语音哈希']
        voice_file = row['语音文件名']
        text = str(row['语音文本'])
        src_path = f"{src}/{voice_hash}.wav"

        if Path(src_path).exists() == True:

            if text != "" or text != None:

                if Path(f"{dst}/{dest_path}/{character}").exists() == False:
                    Path(f"{dst}/{dest_path}/{character}").mkdir(parents=True, exist_ok=True)

                dst_path = f"{dst}/{dest_path}/{character}/{voice_file}.wav"

                if mode == "cp":
                    copy(src_path, dst_path)
                elif mode == "mv":
                    move(src_path, dst_path)
                else:
                    print("Invalid mode！Defaulting to move")
                    move(src_path, dst_path)

                lab_file_path = f"{dst}/{dest_path}/{character}/{voice_file}.lab"
                cleaned_lab = crean_text(text)
                Path(lab_file_path).write_text(cleaned_lab, encoding='utf-8')
        else:
            tqdm.write(f"{voice_hash}.wav does not exist！Skipping...")
    print("Complete！")

sorting_voice(wav_path,"./Data/sorted",args.mode.lower(),language.upper())

#part3
import json
index = Path("./Data/Sorted.json").read_text(encoding="utf-8")
data = json.loads(index)
lab_src = glob(r"./Data/sorted/**/*.lab",recursive=True)

def get_path_by_lang(lang):
    langcodes = ["CHS","EN","JP","KR"]
    path = ['中文 - Chinese', '英语 - English',  '日语 - Japanese', '韩语 - Korean']
    try:
        i = langcodes.index(lang.Upper())
        dest_path = path[i]
    except:
        print("不支持的语言")
        exit()
    return dest_path

path_by_lang = get_path_by_lang(language);
dest = args.final_destination
for lab_file in tqdm(lab_src):
    try:
        src_dir = os.path.dirname(lab_file)
        lab_file_name = os.path.basename(lab_file)
        wav_file_name = lab_file_name.replace(".lab",".wav")
        dst_dir = data[lab_file_name]
        if not os.path.exists(f"{dest}/{path_by_lang}/{dst_dir}"):
            Path(f"{dest}/{path_by_lang}/{dst_dir}").mkdir(parents=True)
        move(f"{src_dir}/{lab_file_name}",f"{dest}/{path_by_lang}/{dst_dir}/{lab_file_name}")
        move(f"{src_dir}/{wav_file_name}",f"{dest}/{path_by_lang}/{dst_dir}/{wav_file_name}")
    except:
        pass

#part4

tags = r'[<>]'

labfiles = glob(f"{final_destination}/**/*.lab", recursive=True)


def check_content(text, regx):
    if re.search(regx, text):
        return True
    else:
        return False


def tag_content(text):
    res = re.findall(r'(<.*?>)', text)
    string = '、'.join(res)
    return string


def get_path_by_lang(lang):
    langcodes = ["CHS", "EN", "JP", "KR"]
    path = ['中文 - Chinese', '英语 - English', '日语 - Japanese', '韩语 - Korean']
    try:
        i = langcodes.index(lang)
        dest_path = path[i]
    except:
        print("不支持的语言")
        exit()
    return dest_path


path_by_lang = get_path_by_lang(language);

for file in tqdm(labfiles):
    try:
        lab_content = Path(file).read_text(encoding='utf-8')
        spk = os.path.basename(os.path.dirname(file))
        lab_file_name = os.path.basename(file)
        wav_file_name = lab_file_name.replace(".lab", ".wav")
        src = f"{final_destination}/{path_by_lang}/{spk}"
        if check_content(lab_content, tags):
            labels = re.sub(r'<.*?>', '', lab_content)
            lab_path = f"{src}/{lab_file_name}"
            Path(lab_path).write_text(labels, encoding='utf-8')
            tqdm.write(f"已清除标注文件 {src}/{lab_file_name} 中的html标签：{tag_content(lab_content)}\n-----------")
    except:
        pass