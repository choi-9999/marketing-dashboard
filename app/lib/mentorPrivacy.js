export function maskPersonName(value) {
  const name = String(value || "").trim();
  const characters = Array.from(name);

  if (characters.length === 0) return "";
  if (characters.length === 1) return "*";
  if (characters.length === 2) return `${characters[0]}*`;

  return `${characters[0]}${"*".repeat(characters.length - 2)}${characters.at(-1)}`;
}

export function sanitizeRawTabsForPublic(payload) {
  if (!payload || !Array.isArray(payload.rawTabs)) return payload;

  return {
    ...payload,
    rawTabs: payload.rawTabs.map((tab) => {
      if (tab?.kind !== "special-mentor" || !Array.isArray(tab.mentorRows)) return tab;

      return {
        ...tab,
        mentorRows: tab.mentorRows.map((row) => ({
          ...row,
          name: maskPersonName(row?.name),
          phone: ""
        }))
      };
    })
  };
}
