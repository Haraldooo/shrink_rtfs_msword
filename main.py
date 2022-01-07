from genericpath import exists
import win32com
import win32com.client as win32
import os
import os.path
import pathlib
import shutil
import sqlite3
import datetime
from time import sleep
from pprint import pprint
import click

def walk(top):
    "uses scandir to find all files"
    dirs, files = [], []
    for entry in os.scandir(top):
        (dirs if entry.is_dir() else files).append(entry.path)
    yield top, dirs, files
    for path in dirs:
        for x in walk(path):
            yield x

def find_rtf_files(top_dir, rtf_list_file, min_file_size = 1048576):
    """searches the filesystem tree for rtf-files larger 1MByte
    and writes then into a sqlite.db
    """
    con = sqlite3.connect(rtf_list_file)
    with con:
        cur = con.cursor()
        for top, dirs, files in walk(top_dir):
            for f in files:
                if f.endswith('rtf'):
                    file_size = os.path.getsize(f)
                    if file_size > min_file_size:
                        print(f"Found {f}: Size: {file_size/1024/1024}MBytes")
                        cur.execute("INSERT INTO `rtf_list` (file, shrunk) VALUES (?,?)", [f,False])

def shrink_files(rtf_list_file, dry_run):
    con = sqlite3.connect(rtf_list_file)
    with con:
        cur = con.cursor()
        cur.execute("SELECT COUNT(*) FROM `rtf_list`")
        row_count, *_ = cur.fetchone()
        pprint(f"row_count: {row_count}")
        cur.execute("SELECT COUNT(*) FROM `rtf_list` WHERE `shrunk` = 1",) # 1 = True
        already_shrunk, *_ = cur.fetchone()
        pprint(f"already shrunken: {already_shrunk}")
        click.echo(f"{already_shrunk}/{row_count} were already shrunken. Progrssing now")
        pending :int = int(row_count) - int(already_shrunk)
        cur.execute("SELECT rowid, file FROM `rtf_list` WHERE `shrunk` = 0") # 0 = False
        remaining_list = cur.fetchall()
        so_far = 0
        for id, item in remaining_list:
            click.echo(f"id {id} :: item:{item}")
            if os.path.exists(item):
                so_far += 1
                if not dry_run:
                    open_and_save_in_word(item)
                    cur.execute(f"UPDATE rtf_list SET shrunk = 1 WHERE rowid = {id}") # 1 = True
                click.echo(f"{so_far}/{pending}")
                con.commit()

def work_logger(func, file, error):
    "very poor man's logging"
    with open('./logger.log', "a+", encoding="utf-8") as f:
        print(f"{datetime.datetime.now()} file: {file} -- error: {error}")
        print(f"{datetime.datetime.now()} file: {file} -- error: {error}", file=f)

def create_rtf_list_file(rtf_list_file):
    con = sqlite3.connect(rtf_list_file)
    with con:
        cur = con.cursor()
        cur.execute("CREATE TABLE rtf_list(file text not null, shrunk bool)")
        con.commit()

def open_and_save_in_word(rtf_file):
    here = pathlib.Path(__file__).parent.resolve()
    orig_file = rtf_file
    tmp_file_name = r'tmp.rtf'
    tmp_file = here.joinpath(tmp_file_name)
    #try: ToDo: This should get some serious improvement.. 
    shutil.copyfile(orig_file, tmp_file)
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.Documents.Open(str(tmp_file))
    word.ActiveDocument.Save()
    word.Application.Quit()
    sleep(0.2)
    shutil.copyfile(tmp_file, orig_file)
    #except:
    #    work_logger("_", rtf_file, "soMe ErRroR")
    sleep(0.5)

@click.command()
@click.option("--top-dir",
    help="Defines the top of the tree.. ")
@click.option("--scan", is_flag=True,
    help="Scans the directory recursively to find RTF files matching set criteria")
@click.option("--shrink", is_flag=True,
    help="Finds rtfs from db and reduces their file size")
@click.option("--dry-run", is_flag=True,
    help="don't actually shrink them")
def main(top_dir, scan, shrink, dry_run):
    rtf_list_file = "./rft_list_file.db"
    if(scan):
        if not top_dir:
            click.echo("You need to specify --top-dir")
            exit(-1)
        if exists(rtf_list_file):
            click.echo("rtf_list_file.db does already exist.")
            exit(-1)
        else:
            create_rtf_list_file(rtf_list_file)
        find_rtf_files(top_dir, rtf_list_file)
    if(shrink):
        if not exists(rtf_list_file):
            click.echo("rtf_list_file.db not found. Please gernerate first with option '--scan'")
            exit(-1)
        else:
            shrink_files(rtf_list_file, dry_run)

if __name__ == "__main__":
    main()