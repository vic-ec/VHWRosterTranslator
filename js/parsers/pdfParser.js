export async function parsePdfRoster(file) {
  const text = await file.text();
  return { source: file.name, text, entries: [] };
}
