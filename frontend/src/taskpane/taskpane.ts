
let chatMemory: { role: "user" | "ai"; message: string }[] = [];

function copyToClipboard(text: string) {
  navigator.clipboard
    .writeText(text)
    .then(() => alert("✅ Berhasil disalin ke clipboard!"))
    .catch((err) => console.error("❌ Gagal menyalin:", err));
}

function csvToArray(csv: string): string[][] {
  const rows = csv.trim().split("\n");
  return rows.map((row) =>
    row
      .split(/,(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)/)
      .map((cell) => cell.trim().replace(/^["']|["']$/g, ""))
  );
}

function markdownToCsv(markdown: string): string {
  const lines = markdown
    .trim()
    .split("\n")
    .filter((line) => !/^(\s*\|[-\s|]+\|?\s*)$/.test(line));
  const csvLines = lines.map((line) =>
    line
      .trim()
      .replace(/^(\|)/, "")
      .replace(/(\|)$/, "")
      .split("|")
      .map((cell) => cell.trim())
      .join(",")
  );
  return csvLines.join("\n");
}

function pasteCsvToExcel(csv: string, givenRange?: Excel.Range) {
  const dataArray = csvToArray(csv);

  if (!dataArray.length || !Array.isArray(dataArray[0])) {
    alert("❌ Format CSV tidak valid atau kosong.");
    return;
  }

  Excel.run(async (context) => {
    const range = givenRange || context.workbook.getSelectedRange();
    const firstCell = range.getCell(0, 0);
    const numRows = dataArray.length;
    const numCols = dataArray[0].length;
    const targetRange = firstCell.getResizedRange(numRows - 1, numCols - 1);

    // 1. Paste data
    targetRange.values = dataArray;

    // 2. Format header
    const headerRange = firstCell.getResizedRange(0, numCols - 1);
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";
    headerRange.format.font.bold = true;

    // 3. Auto-fit columns
    for (let col = 0; col < numCols; col++) {
      const colRange = firstCell.getOffsetRange(0, col).getEntireColumn();
      colRange.format.autofitColumns();
    }

    // 4. Apply borders
    const borderTypes = [
      Excel.BorderIndex.edgeTop,
      Excel.BorderIndex.edgeBottom,
      Excel.BorderIndex.edgeLeft,
      Excel.BorderIndex.edgeRight,
      Excel.BorderIndex.insideVertical,
      Excel.BorderIndex.insideHorizontal,
    ];

    borderTypes.forEach((borderType) => {
      const border = targetRange.format.borders.getItem(borderType);
      border.style = Excel.BorderLineStyle.continuous;
      border.weight = Excel.BorderWeight.thin;
      border.color = "black";
    });

    await context.sync();

    console.log("📋 CSV berhasil ditempel ke Excel dengan format:", dataArray);
    alert(`✅ Data berhasil ditempel (${numRows} baris × ${numCols} kolom)`);
  }).catch((err) => {
    console.error("❌ Gagal menempel ke Excel:", err);
    alert(
      "❌ Gagal menempel ke Excel. Cek hal berikut:\n• Range cukup luas\n• Format CSV valid\n• Sheet tidak terkunci."
    );
  });
}

const previewBtn = document.getElementById("preview-selected");
previewBtn?.addEventListener("click", async () => {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values");
    await context.sync();
    document.getElementById("selected-preview")!.textContent =
      JSON.stringify(range.values) || "[(Tidak ada range yang dipilih)]";
    console.log("📌 Selected Excel Range:", range.values);
  });
});

const sendButton = document.getElementById("send");
sendButton?.addEventListener("click", async () => {
  const input = document.getElementById("input") as HTMLTextAreaElement;
  const userInput = input.value.trim();
  if (!userInput) return;

  let selectedText = "";

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values");
    await context.sync();
    selectedText = JSON.stringify(range.values);
    console.log("📊 Data yang dikirim:", selectedText);
  });

  try {
    const chatHistory = document.getElementById("chat-history");
    if (!chatHistory) throw new Error("Chat history element not found");

    const loadingIndicator = document.getElementById("loading-indicator");
    if (loadingIndicator) loadingIndicator.style.display = "block";

    const aiResponse = await fetch("http://localhost:3001/api/ai", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ command: userInput, sheetData: selectedText, memory: chatMemory }),
    });

    const aiResponsePayload = await aiResponse.json();
    const reply = aiResponsePayload.reply;

    let csvData = "";
    if (/^\s*\|.*\|\s*$/.test(reply)) {
      csvData = markdownToCsv(reply);
    } else if (/^[^|]+,[^|]+/.test(reply)) {
      csvData = reply.trim();
    }

    chatHistory.innerHTML += `
      <div class="user-message">
        🧑 You: ${userInput}
        <button class="copy-btn" data-text="${encodeURIComponent(userInput)}">📋</button>
      </div>
      <div class="ai-message">
        <strong>🤖 AI:</strong>
        <div style="margin-top: 5px; white-space: pre-wrap">${reply}</div>
        ${
          csvData
            ? `
          <button class="copy-csv-btn" data-csv="${encodeURIComponent(csvData)}">📋 Copy CSV</button>
          <button class="paste-csv-btn" data-csv="${encodeURIComponent(csvData)}">📥 Tempel ke Excel</button>
        `
            : ""
        }
      </div>
    `;

    chatMemory.push({ role: "user", message: userInput });
    chatMemory.push({ role: "ai", message: reply });

    document.querySelectorAll(".copy-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const text = decodeURIComponent((btn as HTMLElement).getAttribute("data-text") || "");
        copyToClipboard(text);
      });
    });

    document.querySelectorAll(".copy-csv-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const csv = decodeURIComponent((btn as HTMLElement).getAttribute("data-csv") || "");
        copyToClipboard(csv);
      });
    });

    document.querySelectorAll(".paste-csv-btn").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const csv = decodeURIComponent((btn as HTMLElement).getAttribute("data-csv") || "");
        try {
          await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            await context.sync();
            pasteCsvToExcel(csv, range);
          });
        } catch (error) {
          console.error("Error getting selected range:", error);
          pasteCsvToExcel(csv);
        }
      });
    });

    input.value = "";
    chatHistory.scrollTop = chatHistory.scrollHeight;
  } catch (error) {
    console.error("Full Error Stack:", error);
    const errorDisplay = document.getElementById("error-display");
    if (errorDisplay) {
      errorDisplay.textContent = error instanceof Error ? error.message : String(error);
      errorDisplay.style.display = "block";
    }
  } finally {
    const loadingIndicator = document.getElementById("loading-indicator");
    if (loadingIndicator) loadingIndicator.style.display = "none";
  }
});

const historyBtn = document.getElementById("show-history");
historyBtn?.addEventListener("click", () => {
  const historyText = chatMemory
    .map((entry) => `${entry.role === "user" ? "🧑" : "🤖"} ${entry.message}`)
    .join("\n\n");
  alert("Riwayat Percakapan:\n\n" + historyText);
});
function csvToTsv(csv: string): string {
  return csv
    .split("\n")
    .map((row) =>
      row
        .split(/,(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)/)
        .map((cell) => cell.replace(/^["']|["']$/g, "").replace(/"/g, '""'))
        .join("\t")
    )
    .join("\n");
}
const generateChartBtn = document.getElementById("generate-chart");
generateChartBtn?.addEventListener("click", async () => {
  try {
    await createChartFromSelection(); // default: columnClustered
  } catch (err) {
    console.error("❌ Gagal membuat grafik:", err);
    await showMessage("❌ Gagal membuat grafik.", "error");
  }
});
function showMessage(message: string, type: "success" | "error" | "info" = "info") {
  const box = document.getElementById("status-box");
  if (!box) return alert(message);

  box.textContent = message;
  box.className = `status ${type}`; // misal styling CSS berdasarkan type
  box.style.display = "block";

  setTimeout(() => (box.style.display = "none"), 4000);
}

async function createChartFromSelection(type: Excel.ChartType = Excel.ChartType.columnClustered) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    const chart = sheet.charts.add(type, range, Excel.ChartSeriesBy.columns);
    chart.title.text = "📊 Grafik Otomatis dari AI";
    chart.legend.position = Excel.ChartLegendPosition.right;
    chart.setPosition(range.getOffsetRange(2, 0), range.getOffsetRange(15, 6));
    chart.load("name");
    await context.sync();
    await showMessage(`✅ Grafik '${type}' berhasil dibuat dan disisipkan!`, "success");
  }).catch(async (err) => {
    console.error("❌ Gagal membuat grafik:", err);
    await showMessage("❌ Gagal membuat grafik. Pastikan range valid dan memiliki data.", "error");
  });
}
