async function translate(sourceLang, targetLang) {
  const selectedText = Office.context.mailbox.item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const text = result.value;
      fetch(`https://api-free.deepl.com/v2/translate`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Authorization': 'DeepL-Auth-Key YOUR_API_KEY_HERE'
        },
        body: `text=${encodeURIComponent(text)}&source_lang=${sourceLang}&target_lang=${targetLang}`
      })
      .then(res => res.json())
      .then(data => {
        const translated = data.translations[0].text;
        document.getElementById('output').textContent = translated;
        Office.context.mailbox.item.body.setSelectedDataAsync(translated);
      });
    }
  });
}