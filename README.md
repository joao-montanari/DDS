
# 🦺 DDS Automático por E-mail

Este projeto realiza o **envio automático de DDS (Diálogo Diário de Segurança)** por e-mail utilizando Python e a biblioteca `pywin32` para integração com o Outlook. Cada dia um tema diferente é enviado, com base em um banco de dados JSON.

---

## 📌 Funcionalidades

- Envio automático de e-mails com temas de DDS.
- Conteúdo formatado em HTML com introdução, tópicos e conclusão.
- Leitura cíclica dos temas, garantindo que todos sejam enviados antes de reiniciar.
- Integração com o **Microsoft Outlook** via `win32com`.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3**
- **pywin32** (`win32com.client`)
- **JSON** como banco de dados
- **Arquivo TXT** para controle de envio sequencial

---

## 📁 Estrutura dos Arquivos

- `main.py` – Script principal da automação.
- `database.json` – Banco de dados com os temas de DDS.
- `current_id.txt` – Armazena o ID do último DDS enviado.

---

## 📂 Estrutura do `database.json`

```json
{
  "themes": [
    {
      "id": 0,
      "theme": "Título do Tema",
      "information": {
        "intro": "Introdução do DDS",
        "topics": [
          {
            "title": "Título do Tópico",
            "list": [
              "Item 1",
              "Item 2"
            ]
          }
        ],
        "conclusion": "Mensagem de encerramento"
      },
      "font": "Fonte ou referência"
    }
  ]
}
```

---

## 🚀 Como Usar

1. Instale o pacote `pywin32`:

   ```bash
   pip install pywin32
   ```

2. Configure a lista de destinatários no `emails` do `main.py`:

   ```python
   emails = ";".join(["email1@empresa.com", "email2@empresa.com"])
   ```

3. Insira seus temas no `database.json`, conforme a estrutura acima.

4. Inicialize o `current_id.txt` com `0`.

5. Execute o script:

   ```bash
   python main.py
   ```

---

## ✉️ Exemplo de E-mail Enviado

```html
<h2>Tema do DDS</h2>
<p>Introdução do conteúdo...</p>
<strong>📌 Título do Tópico</strong>
<ul>
  <li>Item 1</li>
  <li>Item 2</li>
</ul>
<p>Mensagem de conclusão...</p>
<p>🔍 Fonte: Fonte utilizada</p>
<p>🙂 Esse e-mail foi enviado de forma automática</p>
```

---

## 🔄 Funcionamento da Rotina

- O script lê o `current_id.txt` para saber qual DDS enviar.
- Após o envio, ele incrementa esse ID (ou zera se for o último da lista).
- Dessa forma, os temas são enviados de forma **rotativa e automática**.

---

## ✅ Requisitos

- Ter o **Microsoft Outlook instalado e configurado**.
- Executar o script em um ambiente com o Outlook aberto (sessão logada).

---

## 👨‍💻 Autor

Desenvolvido por João Vitor Montanari da Silva – visando promover uma **cultura de segurança no ambiente de trabalho** com tecnologia e automação.

---

## 📜 Licença

Este projeto está sob a licença MIT. Sinta-se à vontade para usar, modificar e contribuir!
