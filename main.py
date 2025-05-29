import win32com.client as win32
import json

outlook = win32.Dispatch('outlook.application')
emails = ";".join([])

with open("database.json", "r", encoding="utf-8") as filePhrases:
    database = json.load(filePhrases)

with open("current_id.txt", "r", encoding="utf-8") as fileCurrentID:
    current_id = fileCurrentID.read()

themes = database['themes']
if int(current_id) >= len(themes):
    current_id = 0

theme = themes[int(current_id)]
logo_url = "https://www.somatick.com.br/site/wp-content/uploads/2014/06/cipa-logo.png"

topics_html = ""
for topic in theme['information']['topics']:
    topics_html += f"""<strong style="color: #28a745;">📌 {topic['title']}</strong>"""
    topics_html += f"""<ul style="padding-left: 20px; margin-top: 10px;">"""
    for list in topic['list']:
        topics_html += f"<li>{list}</li>"
    topics_html += f"</ul>"

body_html = f"""
<table width="100%%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="background-color:#ffffff;">
  <tr>
    <td align="center" style="padding: 40px 0; background-color:#ffffff;">
      <table width="600" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff; font-family:Arial, sans-serif;">

        <tr>
          <td align="center" bgcolor="#28a745" style="padding: 20px; color: #ffffff; font-size: 22px; font-weight: bold; border-radius: 8px;">
            {theme['theme']}
          </td>
        </tr>

        <tr><td style="height: 20px;"></td></tr>

        <tr>
          <td style="padding: 0 20px; color: #333333; font-size: 15px; line-height: 1.6;">
            {theme['information']['intro']}
          </td>
        </tr>

        <tr><td style="height: 20px;"></td></tr>

        <tr>
          <td style="padding: 0 30px; font-size: 15px; color: #222;">
            {topics_html}
          </td>
        </tr>

        <tr>
          <td style="padding: 0 30px; font-size: 15px; color: #222;">
            {theme['information']['conclusion']}
          </td>
        </tr>

        <tr>
            <td align="center" style="padding: 20px 30px 40px 30px;">
                <table cellpadding="0" cellspacing="0" border="0" width="100%%">
                    <tr>
                        <td width="100" valign="top" align="left">
                            <img src="{logo_url}" alt="Logo CIPA" width="130" style="display:block;" />
                        </td>
                        <td align="left" style="padding-left: 10px; font-size: 13px; color: #333333; font-family: Arial, sans-serif;">
                            <strong>Responsabilidade da Informação:</strong> CIPAA - Comissão Interna de Prevenção de Acidentes e de Assédio<br>
                            <strong>Destinatários:</strong> Funcionários<br/>
                            <strong>🔍 Fonte:</strong> {theme['font']}<br/>
                            🙂 Esse e-mail foi gerado de forma automática
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
"""

email = outlook.CreateItem(0)
email.To = emails
email.Subject = 'Diálogo Semanal de Segurança - CIPA'
email.HTMLBody = body_html

email.Send()

new_current_id = int(current_id) + 1
with open("current_id.txt", "w", encoding="utf-8") as fileCurrentID:
    fileCurrentID.write(str(new_current_id))