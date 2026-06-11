import json
import os


def main():
    with open("tag_dict.json", "r") as f:
        tag_dict = json.load(f)

    if os.path.exists("plant_tag_dict.json"):
        with open("plant_tag_dict.json",mode = "r") as f:
            plant_tag_dict = json.load(f)
    else:
        plant_tag_dict = {}
    
    for i in tag_dict:
        if i not in plant_tag_dict:
            print (i , tag_dict[i])
            Plant_Area = input("What area of the plant is this refering to(int, 1-9):")
            plant_tag_dict[i] = {"desc":tag_dict[i], "Plant_Area": Plant_Area}
            with open("plant_tag_dict.json", 'w') as f:
                json.dump(plant_tag_dict,f,indent=2, separators=(',',': '))

if __name__ == "__main__":
    main()


