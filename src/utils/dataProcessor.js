const JIRA_PREFIXES = ['SIG', 'TTSH', 'TSH'];
const SD_PREFIXES = ['HELP'];

export function isJiraKey(key) {
  if (!key) return false;
  const k = key.toString().toUpperCase();
  return JIRA_PREFIXES.some((p) => k.startsWith(p + '-'));
}

export function isSDKey(key) {
  if (!key) return false;
  const k = key.toString().toUpperCase();
  return SD_PREFIXES.some((p) => k.startsWith(p + '-'));
}

export function parseVersionSet(versionStr) {
  if (!versionStr || versionStr.toString().trim() === '') return new Set();
  return new Set(
    versionStr
      .toString()
      .split(',')
      .map((v) => v.trim())
      .filter((v) => v !== ''),
  );
}

function parseLinkedKeys(linkedStr) {
  if (!linkedStr || linkedStr.toString().trim() === '') return [];
  return linkedStr
    .toString()
    .split(',')
    .map((v) => v.trim())
    .filter((v) => v !== '');
}

export function processData(jiraRows, sdRows) {
  // Build lookup maps
  const jiraDataMap = new Map();
  for (const row of jiraRows) {
    jiraDataMap.set(row.key, { fixVersions: row.fixVersions, linkedIssues: row.linkedIssues, url: row.url });
  }

  const helpDataMap = new Map();
  for (const row of sdRows) {
    helpDataMap.set(row.key, { fixVersions: row.fixVersions, linkedIssues: row.linkedIssues, url: row.url });
  }

  // Collect unique cross-system pairs as "jiraKey|||helpKey"
  const pairsSet = new Set();

  // Pass 1: from Jira file - Jira tasks linking to HELP tasks
  for (const row of jiraRows) {
    if (!isJiraKey(row.key)) continue;
    for (const linkedKey of parseLinkedKeys(row.linkedIssues)) {
      if (isSDKey(linkedKey)) {
        pairsSet.add(`${row.key}|||${linkedKey}`);
      }
    }
  }

  // Pass 2: from SD file - HELP tasks linking to Jira tasks
  for (const row of sdRows) {
    if (!isSDKey(row.key)) continue;
    for (const linkedKey of parseLinkedKeys(row.linkedIssues)) {
      if (isJiraKey(linkedKey)) {
        pairsSet.add(`${linkedKey}|||${row.key}`);
      }
    }
  }

  const pairedKeys = new Set();
  const matches = [];

  for (const pairKey of pairsSet) {
    const [jiraKey, helpKey] = pairKey.split('|||');
    pairedKeys.add(jiraKey);
    pairedKeys.add(helpKey);

    const jiraInfo = jiraDataMap.get(jiraKey);
    const helpInfo = helpDataMap.get(helpKey);

    const jiraFVStr = jiraInfo ? jiraInfo.fixVersions : null;
    const helpFVStr = helpInfo ? helpInfo.fixVersions : null;

    let matched = false;
    if (jiraInfo && helpInfo) {
      const jiraVersions = parseVersionSet(jiraFVStr);
      const helpVersions = parseVersionSet(helpFVStr);
      for (const v of jiraVersions) {
        if (helpVersions.has(v)) {
          matched = true;
          break;
        }
      }
    }

    matches.push({
      jiraKey,
      helpKey,
      jiraUrl: jiraInfo?.url || null,
      helpUrl: helpInfo?.url || null,
      jiraFixVersions: jiraInfo ? (jiraFVStr || '') : 'Нет данных',
      helpFixVersions: helpInfo ? (helpFVStr || '') : 'Нет данных',
      status: matched ? 'Совпадает' : 'Не совпадает',
      matched,
    });
  }

  // Sort: mismatches first, then matches
  matches.sort((a, b) => {
    if (a.matched === b.matched) return a.jiraKey.localeCompare(b.jiraKey);
    return a.matched ? 1 : -1;
  });

  // Tasks without any cross-system link
  const unlinked = [];

  for (const row of jiraRows) {
    if (!pairedKeys.has(row.key)) {
      unlinked.push({
        key: row.key,
        url: row.url || null,
        system: 'Jira',
        fixVersions: row.fixVersions,
        linkedIssues: row.linkedIssues,
      });
    }
  }

  for (const row of sdRows) {
    if (!pairedKeys.has(row.key)) {
      unlinked.push({
        key: row.key,
        url: row.url || null,
        system: 'Service Desk',
        fixVersions: row.fixVersions,
        linkedIssues: row.linkedIssues,
      });
    }
  }

  return { matches, unlinked, pairedCount: pairsSet.size };
}
