import os
import anthropic

# API key diambil dari environment variable, JANGAN hardcode di kode.
# Set dulu di terminal sebelum run: export ANTHROPIC_API_KEY="key-kamu-yang-baru"
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

response = client.messages.create(
    model="claude-sonnet-4-5",
    max_tokens=100,
    messages=[
        {"role": "user", "content": "Halo, kamu bisa bantu kategorisasi transaksi akuntansi?"}
    ]
)

print(response.content[0].text)