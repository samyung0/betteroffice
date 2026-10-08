// formatSdtDate's month(width) from contentControlValues.ts, in Node (the
// Office worker's runtime), for the parity table in src/docx/sdt.rs.
//   node crates/office-service/parity/months.mjs
const known = (tag) => {
  try {
    return Intl.DateTimeFormat.supportedLocalesOf(tag).length > 0;
  } catch {
    return false;
  }
};
const month = (language, y, m, d, width) => {
  const date = new Date(Date.UTC(y, m - 1, d));
  const locale = language && known(language) ? language : "en";
  return new Intl.DateTimeFormat(locale, { month: width, timeZone: "UTC" }).format(date);
};
const languages = ["en-US", "ja-JP", "zh-TW", "zh-CN", "zh-HK", "ko-KR", "fr-FR", "de-DE", "es-ES", "pt-BR", "it-IT", "ru-RU", "vi-VN", "th-TH", "ar-SA", "he-IL", "fa-IR", "hi-IN", "nl-NL", "pl-PL", "en_US", "", "not a tag", "xx-XX", "und", "qaa", "x-none", "en-ZZ", "zz", "tlh"];
const rows = [];
for (const language of languages)
  for (let m = 1; m <= 12; m++) rows.push([language, m, month(language, 2026, m, 15, "long"), month(language, 2026, m, 15, "short")]);
console.log(JSON.stringify(rows));
console.log(process.versions.node, process.versions.icu, process.versions.cldr, Intl.DateTimeFormat().resolvedOptions().locale);
