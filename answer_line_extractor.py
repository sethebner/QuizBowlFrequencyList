import ast
import configparser
from collections import defaultdict
from docx import Document
import os
from os import path
import subprocess
import sys
from tqdm import tqdm

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

config_file_name = 'config.ini'
if len(sys.argv) > 1:
    config_file_name = sys.argv[1]
CONFIG_FILE = os.path.join(application_path, config_file_name)

def get_documents(my_paths):
    documents = { 'docx' : [], 'doc' : [], 'pdf' : [] }
    for my_path in my_paths:
        try:
            for f in os.listdir(my_path):
                for key in documents.keys():
                    if f.endswith(key):
                        documents[key].append(''.join([my_path, '/', f]))
                        break
        except FileNotFoundError:
            print('Could not find {}'.format(my_path))
            continue
        except Exception as e:
            print(e)
        
    return documents

def parse_docx(documents, frequencies):
    questions = []
    print('Parsing docx files.....')
    for doc in tqdm(documents['docx']):
        document = Document(doc)

        questions = [ p for p in document.paragraphs if 'ANSWER:' in p.text.upper() ]

        separated_questions = [ q.text.split('\n') for q in questions ]

        answer_lines = [ [ question_part for question_part in sq if 'ANSWER:' in question_part ] for sq in separated_questions ]
        answer_lines = [ item for al in answer_lines for item in al ]
        answer_lines = [ ''.join(answer.split('ANSWER:')[1:]).strip() for answer in answer_lines ]
        answer_lines = [ ''.join(answer.split('[')[0]).strip() for answer in answer_lines ]

        for q in answer_lines:
            frequencies[q] += 1

def parse_pdf(documents, frequencies):
    print('Parsing pdf files.....')
    for doc in tqdm(documents['pdf']):
        try:
            subprocess.call('pdftotext "{}"'.format(doc), shell=True)

            with open( doc.replace('.pdf', '.txt'), 'r' ) as f:
                document = f.read()
                lines = document.split('\n')

                answer_lines = [ line for line in lines if line.startswith('ANSWER:') ]
                answer_lines = [ line.split('ANSWER:')[1].strip() for line in answer_lines ]

                for q in answer_lines:
                    frequencies[q] += 1

            subprocess.call('rm "{}"'.format( doc.replace('.pdf', '.txt') ), shell=True)
        except:
            print('Failed to parse {}'.format(doc))
            
def parse_doc(documents, frequencies):
    questions = []
    print('Parsing doc files.....')
    for doc in tqdm(documents['doc']):
        document = subprocess.getoutput('antiword "{}"'.format(doc))

        lines = document.split('\n')

        answer_lines = [ line for line in lines if line.startswith('ANSWER:') ]
        answer_lines = [ line.split('ANSWER:')[1].strip() for line in answer_lines ]

        for q in answer_lines:
            frequencies[q] += 1

def print_results(frequencies):
    output = sorted(frequencies.items(), key=lambda item: item[1])
    for item in output:
        print(item)

def get_packet_dir(config, header):
    return config.get(*header)

def get_packet_list(config, header, packet_dir):
    packet_list = ast.literal_eval(config.get(*header))
    packet_list = list(set(packet_list))
    packet_list = list(map(lambda s: os.path.join(application_path, packet_dir, s), packet_list))
    return packet_list

def main():
    try:
        parsers = { 'docx' : parse_docx,
                    'doc' : parse_doc,
                    'pdf' : parse_pdf }

        config = configparser.SafeConfigParser()

        try:
            config.read(CONFIG_FILE)
            packet_dir = get_packet_dir(config, ['main', 'packet_dir'])
            if not os.path.isdir(os.path.join(application_path, packet_dir)):
                print('Could not find the directory: {}'.format(os.path.join(application_path, packet_dir)))
                raise NotADirectoryError
            
            packet_list = get_packet_list(config, ['main', 'packet_list'], packet_dir)
            
            frequencies = defaultdict(int)

            documents = get_documents(packet_list)

            for parser in parsers.values():
                parser(documents, frequencies)

            print_results(frequencies)
            print('Extracted {} different answer lines from {} questions.'.format(len(frequencies), sum(frequencies.values())))

        except NotADirectoryError:
            pass
        
        except Exception as e:
            print('ERROR: Could not read the specified configuration file.')
            print(e)
            
    except KeyboardInterrupt:
        pass
        
if __name__ == '__main__':
    main()
