function normalizeLineBreaks(text: string): string {
  return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

function toPlainTextForOutlook(text: string): string {
  return normalizeLineBreaks(text).replace(/\n/g, "\r\n");
}

export function textToOutlookHtml(text: string): string {
  return normalizeLineBreaks(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;")
    .replace(/\n/g, "<br>");
}

function setSelectedBodyData(
  body: Office.Body,
  data: string,
  coercionType: Office.CoercionType
): Promise<void> {
  return new Promise((resolve, reject) => {
    body.setSelectedDataAsync(data, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(result.error ?? new Error("Could not insert text into Outlook."));
    });
  });
}

export function insertTextIntoEmail(text: string): Promise<void> {
  const body = typeof Office === "undefined" ? undefined : Office.context?.mailbox?.item?.body;

  if (!body) {
    return Promise.reject(new Error("Outlook compose body is not available."));
  }

  return setSelectedBodyData(body, textToOutlookHtml(text), Office.CoercionType.Html).catch(
    async (error) => {
      console.warn("HTML insertion failed. Falling back to plain text insertion.", error);
      await setSelectedBodyData(body, toPlainTextForOutlook(text), Office.CoercionType.Text);
    }
  );
}

export async function copyTextToClipboard(text: string): Promise<void> {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textArea = document.createElement("textarea");
  textArea.value = text;
  textArea.style.position = "fixed";
  textArea.style.left = "-9999px";
  textArea.style.top = "0";
  document.body.appendChild(textArea);
  textArea.focus();
  textArea.select();

  try {
    const copied = document.execCommand("copy");

    if (!copied) {
      throw new Error("Browser copy command failed.");
    }
  } finally {
    document.body.removeChild(textArea);
  }
}
