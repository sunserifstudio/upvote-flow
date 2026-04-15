function onFormSubmit(e) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName("Form Responses 1");
  const formUrl = ss.getFormUrl();
  const form = FormApp.openByUrl(formUrl);

  if (!responseSheet) throw new Error("Sheet 'Form Responses 1' not found.");
  if (!formUrl) throw new Error("Spreadsheet not linked to a Form.");

  const headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];

  const suggestCol = headers.findIndex(h => h.startsWith("Suggest a New Meeting Topic")) + 1;
  const voteCol = headers.findIndex(h => h.startsWith("Vote for a Meeting Topic")) + 1;

  if (suggestCol === 0 || voteCol === 0) {
    throw new Error("Required columns not found.");
  }

  const lastRow = responseSheet.getLastRow();
  if (lastRow < 2) return;

  const data = responseSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  let suggestionCounts = {};
  let voteCounts = {};

  data.forEach(row => {

    const suggestion = row[suggestCol - 1];
    const votes = row[voteCol - 1];
    const timestamp = row[0];

    const date = new Date(timestamp);
    const monthKey = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM");
    const monthDisplay = Utilities.formatDate(date, Session.getScriptTimeZone(), "M/yyyy");

    if (suggestion && suggestion.toString().trim() !== "") {
      const topic = suggestion.toString().trim();
      suggestionCounts[topic] = (suggestionCounts[topic] || 0) + 1;
    }

    if (votes && votes.toString().trim() !== "") {

      const selected = votes.toString().split(",");

      selected.forEach(v => {

        const cleanTopic = v.replace(/\s*\(\d+\s+vote(s)?\)$/i, "").trim();

        if (!voteCounts[cleanTopic]) voteCounts[cleanTopic] = { total: 0 };

        if (!voteCounts[cleanTopic][monthKey]) {
          voteCounts[cleanTopic][monthKey] = { count: 0, display: monthDisplay };
        }

        voteCounts[cleanTopic][monthKey].count += 1;
        voteCounts[cleanTopic].total += 1;

      });
    }

  });

  const rawTopics = [...new Set([
    ...Object.keys(suggestionCounts),
    ...Object.keys(voteCounts)
  ])];

  const cleanTopics = rawTopics.map(topic =>
    topic.replace(/\s*\(\d+\s+vote(s)?\)$/i, "").trim()
  );

  const checkboxItems = form.getItems(FormApp.ItemType.CHECKBOX);
  const voteQuestion = checkboxItems[0].asCheckboxItem();

  const topicsWithVotes = cleanTopics.map(topic => ({
    topic,
    votes: (voteCounts[topic] && voteCounts[topic].total) || 0
  }));

  topicsWithVotes.sort((a, b) => {
    if (b.votes !== a.votes) return b.votes - a.votes;
    return a.topic.localeCompare(b.topic);
  });

  const sortedDisplayTopics = topicsWithVotes.map(t => {
    const voteLabel = t.votes === 1 ? "1 vote" : `${t.votes} votes`;
    return `${t.topic} (${voteLabel})`;
  });

  voteQuestion.setChoiceValues(sortedDisplayTopics);

  let summarySheet = ss.getSheetByName("Topic Summary");

  if (!summarySheet) {
    summarySheet = ss.insertSheet("Topic Summary");
  } else {
    summarySheet.clear();
  }

  let monthsMap = {};

  cleanTopics.forEach(topic => {

    const counts = voteCounts[topic] || {};

    Object.keys(counts).forEach(k => {
      if (k !== "total") {
        monthsMap[k] = counts[k].display;
      }
    });

  });

  const monthsSorted = Object.keys(monthsMap)
    .sort((a, b) => new Date(b) - new Date(a)); // newest first

  const monthsDisplay = monthsSorted.map(k => monthsMap[k]);

  summarySheet.appendRow(["Topic", "Times Suggested", "Total Upvotes", ...monthsDisplay]);

  cleanTopics.sort().forEach(topic => {

    const counts = voteCounts[topic] || {};

    const row = [
      topic,
      suggestionCounts[topic] || 0,
      counts.total || 0,
      ...monthsSorted.map(m => counts[m] ? counts[m].count : 0)
    ];

    summarySheet.appendRow(row);

  });

}

function sendMonthlyExcel() {

  const recipients = [
    "youremail@genericemailprovider.com"
  ];
  const folderId = "1t7xhatv6RmAaV8z99gWACG3g591inR2S"; // replace with your folder's ID

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileName = ss.getName();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Convert to Excel blob
  const blob = ss.getBlob().setName(fileName + " " + today + ".xlsx");

  // --- Save to Drive folder ---
  const folder = DriveApp.getFolderById(folderId);
  folder.createFile(blob); // creates a new file in the folder

  // --- Send email with attachment ---
  MailApp.sendEmail({
    to: recipients,
    subject: "Monthly Form Responses Export",
    body: "Attached is the latest Excel export of the form responses.",
    attachments: [blob]
  });

}
