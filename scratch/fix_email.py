import re

with open('e:/Cylinder_MVP/scratch/just_b64.txt', 'r') as f:
    b64 = f.read().strip()

with open('e:/Cylinder_MVP/cylinder_full_script.gs', 'r', encoding='utf-8') as f:
    code = f.read()

# Replace img src in HTML body
code = re.sub(
    r'<img src="data:image/png;base64,[^"]+"',
    r'<img src="cid:nobleLogo"',
    code
)

# Replace MailApp.sendEmail call
replacement = f"""  const logoBase64 = "{b64}";
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(logoBase64), "image/png", "nobleLogo");

  try {{
    MailApp.sendEmail({{
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      name: EMAIL_SENDER_NAME,
      inlineImages: {{ nobleLogo: logoBlob }}
    }});"""

code = code.replace("""  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      name: EMAIL_SENDER_NAME
    });""", replacement)

with open('e:/Cylinder_MVP/cylinder_full_script.gs', 'w', encoding='utf-8') as f:
    f.write(code)

print("Replacement successful")
