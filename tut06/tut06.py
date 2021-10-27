
import re
import os
import shutil

os.system('cls') 
if(os.path.isdir(r'.\correct_srt')):
      pass
else:
    os.makedirs('.\correct_srt')
#os.makedirs('.\correct\game of thrones')
def rename(num,seas,epi):
    path=""
    regex=""
    if num==2:
      if(os.path.isdir(r'.\correct_srt\Game of Thrones')):
          shutil.rmtree('.\correct_srt\Game of Thrones')
      shutil.copytree(r".\wrong_srt\Game of Thrones",r".\correct_srt\Game of Thrones",symlinks=False, ignore=None, copy_function=shutil.copy2, ignore_dangling_symlinks=False, dirs_exist_ok=False)
      path = r'.\correct_srt\Game of Thrones'
      regex = r"([a-zA-Z\s]+)- ([0-9]+)x([0-9]+) - ([a-zA-Z\s]+)"
    elif num==1:
       if(os.path.isdir(r'.\correct_srt\Breaking Bad')):
          shutil.rmtree('.\correct_srt\Breaking Bad')
       shutil.copytree(r".\wrong_srt\Breaking Bad",r".\correct_srt\Breaking Bad",symlinks=False, ignore=None, copy_function=shutil.copy2, ignore_dangling_symlinks=False, dirs_exist_ok=False)
       path = r'.\correct_srt\Breaking Bad'
       regex = r"([a-zA-Z\s]+)s([0-9]+)e([0-9]+)"
    elif num==3:
       if(os.path.isdir(r'.\correct_srt\Lucifer')):
          shutil.rmtree('.\correct_srt\Lucifer')
       shutil.copytree(r".\wrong_srt\Lucifer",r".\correct_srt\Lucifer",symlinks=False, ignore=None, copy_function=shutil.copy2, ignore_dangling_symlinks=False, dirs_exist_ok=False)
       path = r'.\correct_srt\Lucifer'
       regex = r"([a-zA-Z\s]+)- ([0-9]+)x([0-9]+) - ([a-zA-Z\s]+)"

    for filename in os.listdir(path):
      extension=""
      lis=filename
      match = re.search(regex,lis)
      ext=re.search(r".srt$",lis)
      if ext:
        extension=".srt"
      else:
        extension=".mp4"
    # storing the episode and season number and also deleting leading zero
      episode=int(match[3])
      episode=str(episode)
      
      season=int(match[2])
      season=str(season)
    # now making the season and eppisode number equal to their corresponding padding
      while seas>len(season):
        season='0'+season
      while epi>len(episode):
        episode='0'+episode

      path1=path+'/'+filename
      if(num==1):
        name=match[1]+"-"+" " + "season" +" "+season+" " + "episode"+ " "+episode + extension  
      else:
        name=match[1]+"-"+" " + "season" +" "+season+" " + "episode"+ " "+episode +" "+ "-"+" " + match[4]+extension 
      new=path+'/'+name
      os.rename(path1,new)
    

def regex_renamer():

	# Taking input from the user
   
 print("1. Breaking Bad")
 print("2. Game of Thrones")
 print("3. Lucifer") 
 webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))
 if webseries_num>3 or webseries_num<0:
       print("Please enter a valid webseries num ")
       regex_renamer()
 season_padding = int(input("Enter the Season Number Padding: "))
 episode_padding = int(input("Enter the Episode Number Padding: "))
 rename(webseries_num,season_padding,episode_padding)
 
   
regex_renamer()