import markdown
import sys

def main():
    paths = sys.argv
    for i in paths[1:]:
        with open( i, 'r') as f:
            text = f.read()
            html = markdown.markdown(text)
        with open( i[:-3] + ".html" , 'w') as f:
            f.write(html) 
    return

main()