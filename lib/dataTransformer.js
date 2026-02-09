// Data transformer for vocabulary Excel sheet
// Transforms flat Excel data into structured format grouped by topics

/**
 * Parse Excel data and structure it by topics
 * @param {Object} excelData - Raw data from Microsoft Graph API
 * @returns {Object} Structured data with topics and vocabulary words
 */
function transformVocabData(excelData) {
  if (!excelData || !excelData.worksheets || excelData.worksheets.length === 0) {
    return excelData;
  }

  const result = {
    fileName: excelData.fileName,
    fileSize: excelData.fileSize,
    worksheets: []
  };

  // Process each worksheet
  excelData.worksheets.forEach(worksheet => {
    if (worksheet.error || !worksheet.values || worksheet.values.length === 0) {
      result.worksheets.push(worksheet);
      return;
    }

    const structuredWorksheet = {
      name: worksheet.name,
      range: worksheet.range,
      rowCount: worksheet.rowCount,
      columnCount: worksheet.columnCount,
      topics: parseTopics(worksheet.values),
      statistics: null // Will be calculated
    };

    // Calculate statistics
    structuredWorksheet.statistics = calculateStatistics(structuredWorksheet.topics);
    
    // Calculate statistics for each topic
    if (structuredWorksheet.topics && structuredWorksheet.topics.length > 0) {
      structuredWorksheet.topics.forEach(topic => {
        topic.statistics = calculateTopicStatistics([topic]);
      });
    }
    
    result.worksheets.push(structuredWorksheet);
  });

  // Add cached metadata if present
  if (excelData._cached !== undefined) {
    result._cached = excelData._cached;
  }
  if (excelData._cachedAt) {
    result._cachedAt = excelData._cachedAt;
  }
  if (excelData._fetchedAt) {
    result._fetchedAt = excelData._fetchedAt;
  }

  return result;
}

/**
 * Parse rows into topics with vocabulary words
 * @param {Array} rows - 2D array of cell values
 * @returns {Array} Array of topics with their words
 */
function parseTopics(rows) {
  const topics = [];
  let currentTopic = null;

  // Skip header row if it exists (check if first row looks like a header)
  const startRow = isHeaderRow(rows[0]) ? 1 : 0;

  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    
    // Skip empty rows
    if (isEmptyRow(row)) continue;

    const topicName = getCell(row, 1); // Column B (index 1)
    
    // Check if this row starts a new topic
    if (topicName && topicName.trim() !== '') {
      // Save previous topic if exists
      if (currentTopic) {
        topics.push(currentTopic);
      }

      // Start new topic
      currentTopic = {
        name: topicName.trim(),
        startRow: i + 1, // 1-based row number
        words: []
      };

      // The first row of a topic may also contain a word
      const word = parseWord(row, i + 1);
      if (word && word.word) {
        currentTopic.words.push(word);
      }
    } else if (currentTopic) {
      // This row belongs to the current topic
      const word = parseWord(row, i + 1);
      if (word && word.word) {
        currentTopic.words.push(word);
      }
    } else {
      // No current topic - this might be a word without a topic
      // Create a default "Uncategorized" topic
      if (!currentTopic) {
        currentTopic = {
          name: 'Uncategorized',
          startRow: i + 1,
          words: []
        };
      }
      const word = parseWord(row, i + 1);
      if (word && word.word) {
        currentTopic.words.push(word);
      }
    }
  }

  // Add the last topic
  if (currentTopic) {
    topics.push(currentTopic);
  }

  return topics;
}

/**
 * Parse a single word from a row
 * @param {Array} row - Array of cell values
 * @param {number} rowNumber - 1-based row number
 * @returns {Object} Word object
 */
function parseWord(row, rowNumber) {
  const word = getCell(row, 3); // Column D (index 3)
  
  // Skip if no word
  if (!word || word.trim() === '') {
    return null;
  }

  return {
    rowNumber: rowNumber,
    order: getCell(row, 0), // Column A
    flag: getCell(row, 2), // Column C
    word: word.trim(), // Column D
    partOfSpeech: getCell(row, 4), // Column E
    pronunciation: getCell(row, 5), // Column F
    meaning: getCell(row, 6), // Column G
    exampleSentence: getCell(row, 7), // Column H
    synonyms: getCell(row, 8), // Column I
    dayOfWeek: getCell(row, 9), // Column J
    date: getCell(row, 10) // Column K
  };
}

/**
 * Calculate statistics for topics
 * @param {Array} topics - Array of topics
 * @returns {Object} Statistics object
 */
function calculateStatistics(topics) {
  const stats = {
    totalTopics: topics.length,
    totalWords: 0,
    byFlag: {
      new: 0, // N or n
      known: 0, // Y or y
      forgotten: 0, // ?
      learned: 0 // ok
    }
  };

  topics.forEach(topic => {
    stats.totalWords += topic.words.length;
    
    topic.words.forEach(word => {
      const flag = (word.flag || '').toString().toLowerCase().trim();
      if (flag === 'n') {
        stats.byFlag.new++;
      } else if (flag === 'y') {
        stats.byFlag.known++;
      } else if (flag === '?') {
        stats.byFlag.forgotten++;
      } else if (flag === 'ok') {
        stats.byFlag.learned++;
      }
    });
  });

  return stats;
}

/**
 * Calculate statistics for a single topic
 * @param {Array} topics - Array with one topic
 * @returns {Object} Statistics object for the topic
 */
function calculateTopicStatistics(topics) {
  const stats = {
    totalWords: 0,
    byFlag: {
      new: 0,
      known: 0,
      forgotten: 0,
      learned: 0
    }
  };

  topics.forEach(topic => {
    stats.totalWords += topic.words.length;
    
    topic.words.forEach(word => {
      const flag = (word.flag || '').toString().toLowerCase().trim();
      if (flag === 'n') {
        stats.byFlag.new++;
      } else if (flag === 'y') {
        stats.byFlag.known++;
      } else if (flag === '?') {
        stats.byFlag.forgotten++;
      } else if (flag === 'ok') {
        stats.byFlag.learned++;
      }
    });
  });

  return stats;
}

/**
 * Get cell value safely
 * @param {Array} row - Row array
 * @param {number} index - Column index
 * @returns {string} Cell value or empty string
 */
function getCell(row, index) {
  if (!row || index >= row.length) return '';
  const value = row[index];
  return value !== null && value !== undefined ? String(value).trim() : '';
}

/**
 * Check if a row is empty
 * @param {Array} row - Row array
 * @returns {boolean} True if row is empty
 */
function isEmptyRow(row) {
  if (!row || row.length === 0) return true;
  return row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
}

/**
 * Check if a row looks like a header
 * @param {Array} row - Row array
 * @returns {boolean} True if row appears to be a header
 */
function isHeaderRow(row) {
  if (!row || row.length === 0) return false;
  const firstCell = String(getCell(row, 0)).toLowerCase();
  const secondCell = String(getCell(row, 1)).toLowerCase();
  return firstCell === 'order' || secondCell === 'topic' || secondCell === 'session';
}

module.exports = {
  transformVocabData,
  parseTopics,
  parseWord,
  calculateStatistics,
  calculateTopicStatistics
};
