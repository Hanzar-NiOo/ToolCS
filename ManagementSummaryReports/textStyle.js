// ---- Shared style definitions ----
const TEXT_STYLES = {
  title:    { fontFamily: "Calibri",       fontSize: 30, fontColor: "#000000" },
  subtitle: { fontFamily: "Calibri",       fontSize: 18, fontColor: "#000000" },
  heading1: { fontFamily: "Calibri Light", fontSize: 16, fontColor: "#000000" },
  heading2: { fontFamily: "Calibri Light", fontSize: 14, fontColor: "#000000" },
  heading3: { fontFamily: "Calibri Light", fontSize: 12, fontColor: "#000000" },
  body:     { fontFamily: "Calibri",       fontSize: 11, fontColor: "#000000" }
};

function applyTextStyle(range, styleName, options)
{
  const style = TEXT_STYLES[styleName];
  if (!style)
  {
    throw new Error("Unknown text style: " + styleName);
  }

  range.setFontFamily(style.fontFamily)
       .setFontSize(style.fontSize)
       .setFontColor(style.fontColor);

  if (options && options.bold)
  {
    range.setFontWeight("bold");
  }
}