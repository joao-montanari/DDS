import win32com.client as win32
import json

outlook = win32.Dispatch('outlook.application')
emails = ";".join([])

with open("database.json", "r", encoding="utf-8") as filePhrases:
    database = json.load(filePhrases)

with open("current_id.txt", "r", encoding="utf-8") as fileCurrentID:
    current_id = fileCurrentID.read()

themes = database['themes']
theme = themes[int(current_id)]

topics_html = ""
for topic in theme['information']['topics']:
    topics_html += f"<strong>📌 {topic['title']}</strong>"
    topics_html += f"<ul>"
    for list in topic['list']:
        topics_html += f"<li>{list}</li>"
    topics_html += f"</ul>"

email = outlook.CreateItem(0)
email.To = emails
email.Subject = 'Diálogo Semanal de Segurança - CIPA'
email.HTMLBody = f"""
<h2>{theme['theme']}</h2>
<p>{theme['information']['intro']}</p>
{topics_html}
<p>{theme['information']['conclusion']}</p>
<p>🔍 Fonte: {theme['font']}</p>
<p>🙂 Esse e-mail foi enviado de forma automática</p>
"""

email.Send()

if int(current_id) >= len(themes):
    new_current_id = 0
else:
    new_current_id = int(current_id) + 1

with open("current_id.txt", "w", encoding="utf-8") as fileCurrentID:
    fileCurrentID.write(str(new_current_id))