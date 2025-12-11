const CONFIG = {
  default_days: 7,
  max_days: 365,
  require_impressions: true,
  include_pmax: true,
  include_search: true,
  automate_exclusions: false,
  fuzzy_threshold: 0.9,
  similarity_min_threshold: 0.05,
  spreadsheet_url: "URL here",
  branded_terms: [
    'nx', 'nxx', 'hyperlynx', 'tia portal', 'simcenter', 'teamcenter', 'heeds', 'flotherm', 'simatic', 'tecnomatix', 'solid edge', 'capital',
    'aprisa', 'calibre', 'tessent', 'opcenter', 'nx mach', 'nastran', 'star ccm', 'polarion', 'healthineers', 'dotmatics', 'altair', 'geolus',
    'jt open', 'flomaster', 'testlab', 'mentor graphics', 'unigraphics', 'empresa', 'microred', 'amesim', 'culgi', 'vesys'
  ]
};

function main() {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(CONFIG.spreadsheet_url);
    const days = Math.min(Math.max(CONFIG.default_days, 1), CONFIG.max_days);
    const date_period = build_date_period(days);
    
    create_branded_terms_sheet(spreadsheet);
    
    let all_branded_terms = [];
    let account_info = null;
    
    if (CONFIG.include_pmax) {
      const pmax_query = build_pmax_query(date_period);
      const result = run_enhanced_report(pmax_query, spreadsheet, days, 'PMAX');
      all_branded_terms = all_branded_terms.concat(result.branded_terms);
      if (!account_info) account_info = result.account_info;
    }
    
    if (CONFIG.include_search) {
      const search_query = build_search_query(date_period);
      const result = run_enhanced_report(search_query, spreadsheet, days, 'Search');
      all_branded_terms = all_branded_terms.concat(result.branded_terms);
      if (!account_info) account_info = result.account_info;
    }
    
    if (CONFIG.automate_exclusions && all_branded_terms.length > 0 && account_info) {
      const unique_branded = [...new Set(all_branded_terms)];
      add_to_exclusion_list(account_info.customer_name, unique_branded);
    }
    
    Logger.log("Search terms analysis completed successfully");
  } catch (error) {
    Logger.log(`Error in main: ${error.message}`);
    throw error;
  }
}

function build_date_period(days) {
  if (days === 7) return "DURING LAST_7_DAYS";
  if (days === 14) return "DURING LAST_14_DAYS";
  if (days === 30) return "DURING LAST_30_DAYS";
  
  const end_date = new Date();
  const start_date = new Date();
  start_date.setDate(end_date.getDate() - days);
  
  return `BETWEEN '${formatDate(start_date)}' AND '${formatDate(end_date)}'`;
}

function build_pmax_query(date_period) {
  let query = `SELECT 
    campaign_search_term_view.search_term,
    metrics.cost_micros,
    metrics.impressions,
    metrics.clicks,
    metrics.conversions,
    campaign.advertising_channel_type,
    campaign_search_term_view.resource_name,
    segments.search_term_targeting_status,
    campaign.id,
    campaign.name,
    customer.id,
    customer.descriptive_name
    FROM campaign_search_term_view
    WHERE segments.date ${date_period}
    AND campaign.advertising_channel_type IN ('PERFORMANCE_MAX')
    AND segments.search_term_targeting_status IN ('NONE', 'UNKNOWN')`;
  
  if (CONFIG.require_impressions) query += " AND metrics.impressions > 0";
  query += " ORDER BY metrics.cost_micros DESC";
  
  return query;
}

function build_search_query(date_period) {
  let query = `SELECT
    customer.id,
    customer.descriptive_name,
    campaign.id,
    campaign.name,
    campaign.advertising_channel_type,
    search_term_view.search_term,
    metrics.cost_micros,
    metrics.impressions,
    metrics.clicks,
    metrics.conversions,
    search_term_view.status
    FROM search_term_view
    WHERE segments.date ${date_period}
    AND search_term_view.status IN ('NONE', 'UNKNOWN')`;
  
  if (CONFIG.require_impressions) query += " AND metrics.impressions > 0";
  query += " ORDER BY metrics.cost_micros DESC";
  
  return query;
}

function calculate_next_refresh() {
  const now = new Date();
  const next_refresh = new Date(now);
  next_refresh.setDate(now.getDate() + 7);
  next_refresh.setHours(5, 0, 0, 0);
  
  const is_dst = isDaylightSavingTime(next_refresh);
  const year = next_refresh.getFullYear();
  const month = (next_refresh.getMonth() + 1).toString().padStart(2, '0');
  const day = next_refresh.getDate().toString().padStart(2, '0');
  
  return `${year}-${month}-${day} 05:00:00 ${is_dst ? 'EDT' : 'EST'}`;
}

function create_branded_terms_sheet(spreadsheet) {
  try {
    let branded_sheet = spreadsheet.getSheetByName('branded_terms_reference');
    if (!branded_sheet) {
      branded_sheet = spreadsheet.insertSheet('branded_terms_reference');
    }
  
    branded_sheet.clear();
    branded_sheet.getRange(1, 1).setValue('Branded Terms Reference List');
    branded_sheet.getRange(1, 1).setFontWeight('bold');
  
    const terms_data = CONFIG.branded_terms.map(term => [term]);
    if (terms_data.length > 0) {
      branded_sheet.getRange(2, 1, terms_data.length, 1).setValues(terms_data);
    }
  
    branded_sheet.hideSheet();
    Logger.log("Branded terms sheet created");
  } catch (error) {
    Logger.log(`Error creating branded terms sheet: ${error.message}`);
  }
}

function run_enhanced_report(query, spreadsheet, days, report_type) {
  try {
    const report = AdsApp.search(query, { apiVersion: 'v21' });
  
    const campaigns_data = {};
    const campaign_totals = {};
    let row_count = 0;
    let account_info = null;
    const branded_terms = [];
    
    while (report.hasNext()) {
      const row = report.next();
      row_count++;
      
      let search_term, cost, impressions, clicks, conversions, campaign_id, campaign_name, customer_id, customer_name;
      
      if (report_type === 'PMAX') {
        search_term = row.campaignSearchTermView.searchTerm;
        cost = row.metrics.costMicros / 1000000;
        impressions = row.metrics.impressions;
        clicks = row.metrics.clicks;
        conversions = row.metrics.conversions;
        campaign_id = row.campaign.id;
        campaign_name = row.campaign.name;
        customer_id = row.customer.id;
        customer_name = row.customer.descriptiveName;
      } else {
        search_term = row.searchTermView.searchTerm;
        cost = row.metrics.costMicros / 1000000;
        impressions = row.metrics.impressions;
        clicks = row.metrics.clicks;
        conversions = row.metrics.conversions;
        campaign_id = row.campaign.id;
        campaign_name = row.campaign.name;
        customer_id = row.customer.id;
        customer_name = row.customer.descriptiveName;
      }
      
      if (!account_info) {
        account_info = { customer_id, customer_name };
      }
      
      const campaign_key = `${report_type}_${campaign_id}`;
      if (!campaigns_data[campaign_key]) {
        campaigns_data[campaign_key] = {
          campaign_name,
          campaign_id,
          campaign_type: report_type,
          search_terms: []
        };
        campaign_totals[campaign_key] = 0;
      }
      
      campaigns_data[campaign_key].search_terms.push({
        search_term, cost, impressions, clicks, conversions
      });
      
      campaign_totals[campaign_key] += cost;
    }
    
    for (const campaign_key in campaigns_data) {
      const campaign_data = campaigns_data[campaign_key];
      const campaign_total = campaign_totals[campaign_key];
      const sheet_name = `${campaign_data.campaign_type}_${campaign_data.campaign_id}`;
      const campaign_sheet = get_or_create_sheet(spreadsheet, sheet_name);
      
      clear_sheet_content(campaign_sheet);
      add_report_header(campaign_sheet, days, account_info, campaign_data.campaign_name, campaign_data.campaign_id, campaign_data.campaign_type);
      
      const processed_data = [];
      for (const term_data of campaign_data.search_terms) {
        const { search_term, cost, impressions, clicks, conversions } = term_data;
        
        const cpc = clicks > 0 ? cost / clicks : 0;
        const cost_percentage = campaign_total > 0 ? (cost / campaign_total) : 0;
        const max_string_similarity = BRANDED_FUZZY_3(search_term, CONFIG.branded_terms);
        const best_match = find_best_branded_match(search_term);
        const terms_detected = find_detected_branded_terms(search_term);
        const siemens_branded = determine_siemens_branded(search_term, max_string_similarity, terms_detected);
        const exact_match = `[${search_term}]`;
        const phrase_match = `"${search_term}"`;
        
        if (siemens_branded === 'branded') {
          branded_terms.push(search_term);
        }
        
        processed_data.push([
          campaign_data.campaign_id, campaign_data.campaign_name, search_term, best_match, 
          Math.round(max_string_similarity * 1000) / 1000, terms_detected, siemens_branded,
          cost, cost_percentage, impressions, clicks, cpc, conversions, exact_match, phrase_match
        ]);
      }
      
      const header_row = 11;
      
      let headers = ["Campaign ID", "Campaign Name", "Search Term", "Best Match", "Max String Similarity", "Terms Detected", "Siemens Branded", "Cost", "% of search cost", "Impressions", "Clicks", "CPC", "Conversions", "Exact Match", "Phrase Match"];
      
      campaign_sheet.getRange(header_row, 1, 1, headers.length).setValues([headers]);
      
      const max_sim_cell = campaign_sheet.getRange(header_row, 5);
      max_sim_cell.setFormula('=HYPERLINK("https://en.wikipedia.org/wiki/Jaro%E2%80%93Winkler_distance", "Max String Similarity")');
      
      const branded_col_cell = campaign_sheet.getRange(header_row, 7);
      branded_col_cell.setNote('Branded if: detected branded term(s) within search term OR max string similarity >= fuzzy threshold');
      
      if (processed_data.length > 0) {
        campaign_sheet.getRange(header_row + 1, 1, processed_data.length, headers.length).setValues(processed_data);
        format_results_sheet(campaign_sheet, headers.length, header_row, campaign_data.campaign_type);
      }
      
      Logger.log(`Created ${campaign_data.campaign_type} sheet for campaign ${campaign_data.campaign_id}`);
    }
    
    Logger.log(`Processed ${row_count} ${report_type} rows`);
    
    return { branded_terms, account_info };
  } catch (error) {
    Logger.log(`Error in run_enhanced_report: ${error.message}`);
    throw error;
  }
}

function add_report_header(sheet, days, account_info, campaign_name, campaign_id, campaign_type) {
  const current_time = getEasternTime();
  const next_refresh = calculate_next_refresh();
  const end_date = new Date();
  const start_date = new Date();
  start_date.setDate(end_date.getDate() - days);
  
  const date_range = `${formatDate(start_date)} to ${formatDate(end_date)}`;
  const account_display = `${account_info.customer_name} - ${account_info.customer_id}`;
  const campaign_display = `${campaign_name} - ${campaign_id}`;
  const title = campaign_type === 'PMAX' ? 'PMAX Branded Search Term Finder' : 'Search Branded Search Term Finder';
  
  sheet.getRange(1, 1).setValue(title).setFontSize(14).setFontWeight('bold');
  sheet.getRange(2, 1).setValue(`Account: ${account_display}`);
  sheet.getRange(3, 1).setValue(`Campaign: ${campaign_display} (${campaign_type})`);
  sheet.getRange(4, 1).setValue(`Report Period: ${date_range} (${days} days)`);
  sheet.getRange(5, 1).setValue(`Last Refresh: ${current_time}`);
  sheet.getRange(6, 1).setValue(`Next Refresh: ${next_refresh}`);
  sheet.getRange(7, 1).setValue(`Fuzzy Threshold: ${CONFIG.fuzzy_threshold} | Min Similarity: ${CONFIG.similarity_min_threshold}`);
  sheet.getRange(8, 1).setValue(`Total Branded Terms Referenced: ${CONFIG.branded_terms.length}`);
  sheet.getRange(9, 1).setValue('Note: Non-Latin search terms (Cyrillic, Chinese, Japanese, Korean, Arabic, etc.) are not evaluated for fuzzy matching due to character encoding differences, phonetic variations, and algorithm limitations with non-Latin scripts.');
  
  sheet.getRange(2, 1, 7, 1).setFontSize(10);
  sheet.getRange(9, 1).setFontSize(8).setFontStyle('italic');
}

function determine_siemens_branded(search_term, max_string_similarity, terms_detected) {
  if (!search_term || typeof search_term !== 'string') {
    return 'unable to evaluate';
  }
  
  if (is_entirely_non_latin(search_term)) {
    return 'unable to evaluate';
  }
  
  const has_detected_terms = terms_detected.length > 0;
  const meets_fuzzy_threshold = max_string_similarity >= CONFIG.fuzzy_threshold;
  
  if (has_detected_terms || meets_fuzzy_threshold) {
    return 'branded';
  }
  
  return 'not branded';
}

function find_best_branded_match(search_term) {
  if (!search_term || typeof search_term !== 'string') return '';
  
  if (is_entirely_non_latin(search_term)) return '';
  
  const search_lower = search_term.toLowerCase();
  let best_match = '';
  let max_similarity = 0;
  
  for (const branded_term of CONFIG.branded_terms) {
    const branded_lower = branded_term.toLowerCase();
    
    if (search_lower === branded_lower) {
      return branded_term;
    }
    
    const similarity = jaroWinklerSimilarity(search_lower, branded_lower);
    if (similarity > max_similarity) {
      max_similarity = similarity;
      best_match = branded_term;
    }
  }
  
  return max_similarity >= CONFIG.similarity_min_threshold ? best_match : '';
}

function find_detected_branded_terms(search_term) {
  if (!search_term || typeof search_term !== 'string') return '';
  
  if (is_entirely_non_latin(search_term)) return '';
  
  const detected_terms = [];
  
  for (const branded_term of CONFIG.branded_terms) {
    if (is_branded_term_detected(search_term, branded_term)) {
      detected_terms.push(branded_term);
    }
  }
  
  return detected_terms.length > 0 ? detected_terms.join(', ') : '';
}

function is_branded_term_detected(search_term, branded_term) {
  const search_lower = search_term.toLowerCase();
  const branded_lower = branded_term.toLowerCase();
  
  if (search_lower === branded_lower) return true;
  
  const branded_words = branded_lower.split(/\s+/);
  
  if (branded_words.length === 1) {
    const word_boundary_pattern = new RegExp(`\\b${escape_regex(branded_lower)}\\b`);
    return word_boundary_pattern.test(search_lower);
  } else {
    const branded_phrase = branded_words.join('\\s+');
    const phrase_pattern = new RegExp(`\\b${branded_phrase}\\b`);
    return phrase_pattern.test(search_lower);
  }
}

function is_entirely_non_latin(search_term) {
  const cleaned_term = search_term.replace(/[\s\-_.]/g, '');
  if (cleaned_term.length === 0) return false;
  const non_latin_regex = /^[\u0400-\u04FF\u0500-\u052F\u2DE0-\u2DFF\uA640-\uA69F\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\u0600-\u06FF]+$/;
  return non_latin_regex.test(cleaned_term);
}

function escape_regex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function get_or_create_sheet(spreadsheet, sheet_name) {
  let sheet = spreadsheet.getSheetByName(sheet_name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheet_name);
  }
  return sheet;
}

function clear_sheet_content(sheet) {
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).clear();
  }
}

function format_results_sheet(sheet, column_count, header_row, campaign_type) {
  const column_widths = { 1: 110, 2: 210, 3: 260, 4: 140, 5: 145, 6: 160, 7: 130, 8: 90, 9: 130, 10: 90, 11: 70, 12: 80, 13: 90, 14: 150, 15: 150 };
  
  for (let col = 1; col <= column_count; col++) {
    if (column_widths[col]) {
      sheet.setColumnWidth(col, column_widths[col]);
    }
  }
  
  const header_range = sheet.getRange(header_row, 1, 1, column_count);
  header_range.setFontWeight('bold').setBackground('#d9d9d9').setBorder(false, false, true, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID).setWrap(true);
  
  if (sheet.getLastRow() > header_row) {
    const last_row = sheet.getLastRow();
    const data_rows = last_row - header_row;
    
    sheet.getRange(header_row + 1, 5, data_rows, 1).setNumberFormat('0.000');
    sheet.getRange(header_row + 1, 8, data_rows, 1).setNumberFormat('$#,##0.00');
    sheet.getRange(header_row + 1, 9, data_rows, 1).setNumberFormat('0.00%');
    sheet.getRange(header_row + 1, 10, data_rows, 1).setNumberFormat('#,##0');
    sheet.getRange(header_row + 1, 11, data_rows, 1).setNumberFormat('#,##0');
    sheet.getRange(header_row + 1, 12, data_rows, 1).setNumberFormat('$#,##0.00');
    sheet.getRange(header_row + 1, 13, data_rows, 1).setNumberFormat('0.00');
  }
}

function add_to_exclusion_list(account_name, branded_terms) {
  try {
    const exclusion_list_name = `${account_name} - branded exclusions`;
    let keywords_added = 0;
    
    const campaigns_iterator = AdsApp.campaigns().get();
    
    while (campaigns_iterator.hasNext()) {
      const campaign = campaigns_iterator.next();
      
      for (const term of branded_terms) {
        const exact_match_keyword = `[${term}]`;
        campaign.createNegativeKeyword(exact_match_keyword);
        keywords_added++;
      }
    }
    
    Logger.log(`Added ${keywords_added} negative keywords to campaigns for exclusion list: ${exclusion_list_name}`);
  } catch (error) {
    Logger.log(`Error adding to exclusion list: ${error.message}`);
  }
}

function jaroWinklerSimilarity(s1, s2) {
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;
  
  const len1 = s1.length;
  const len2 = s2.length;
  const match_window = Math.floor(Math.max(Math.max(len1, len2) / 2 - 1, 0));
  
  const matches1 = new Array(len1).fill(false);
  const matches2 = new Array(len2).fill(false);
  let matching_chars = 0;
  
  for (let i = 0; i < len1; i++) {
    const start = Math.max(0, i - match_window);
    const end = Math.min(i + match_window + 1, len2);
    
    for (let j = start; j < end; j++) {
      if (!matches2[j] && s1[i] === s2[j]) {
        matches1[i] = true;
        matches2[j] = true;
        matching_chars++;
        break;
      }
    }
  }
  
  if (matching_chars === 0) return 0;
  
  let transpositions = 0;
  let k = 0;
  
  for (let i = 0; i < len1; i++) {
    if (matches1[i]) {
      while (!matches2[k]) k++;
      if (s1[i] !== s2[k]) transpositions++;
      k++;
    }
  }
  
  const m = matching_chars;
  const t = transpositions / 2;
  const jaro = (m / len1 + m / len2 + (m - t) / m) / 3;
  
  let l = 0;
  const max_prefix = Math.min(4, len1, len2);
  for (let i = 0; i < max_prefix; i++) {
    if (s1[i] === s2[i]) {
      l++;
    } else {
      break;
    }
  }
  
  const p = 0.1;
  return jaro + (l * p * (1 - jaro));
}

function BRANDED_FUZZY_3(searchTerm, brandedTerms) {
  if (!searchTerm || !brandedTerms) return 0;
  
  const search_term_str = String(searchTerm);
  
  if (/[\u0400-\u04FF\u0500-\u052F\u2DE0-\u2DFF\uA640-\uA69F\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\u0600-\u06FF]/.test(search_term_str)) {
    return 0;
  }
  
  const search_term_lower = search_term_str.toLowerCase();
  const search_term_words = search_term_lower.split(/\s+/);
  const search_term_words_set = new Set(search_term_words);
  
  let max_similarity = 0;
  
  for (const branded_term of brandedTerms) {
    if (!branded_term) continue;
    
    const branded_term_clean = String(branded_term).toLowerCase();
    if (branded_term_clean === "") continue;
    
    if (branded_term_clean === search_term_lower) return 1;
    
    if (search_term_lower.includes(branded_term_clean)) {
      const substring_score = branded_term_clean.length / search_term_lower.length;
      max_similarity = Math.max(max_similarity, substring_score);
      if (substring_score > 0.9) continue;
    }
    
    const branded_term_words = branded_term_clean.split(/\s+/);
    let has_exact_word_match = false;
    
    for (const word of branded_term_words) {
      if (word.length > 3 && search_term_words_set.has(word)) {
        max_similarity = Math.max(max_similarity, 0.85);
        has_exact_word_match = true;
        break;
      }
    }
    
    if (has_exact_word_match && max_similarity > 0.8) continue;
    
    if (branded_term_clean.length <= 3) {
      const word_boundary_pattern = new RegExp(`\\b${escape_regex(branded_term_clean)}\\b`);
      if (word_boundary_pattern.test(search_term_lower)) {
        max_similarity = Math.max(max_similarity, 0.9);
      }
    } else if (max_similarity < 0.9) {
      const similarity = jaroWinklerSimilarity(search_term_lower, branded_term_clean);
      max_similarity = Math.max(max_similarity, similarity);
    }
  }
  
  return max_similarity;
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getEasternTime() {
  const now = new Date();
  const est_offset = -5.0;
  const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
  let est = new Date(utc + (3600000 * est_offset));
  
  const is_dst = isDaylightSavingTime(now);
  if (is_dst) {
    est = new Date(utc + (3600000 * (est_offset + 1)));
  }
  
  const year = est.getFullYear();
  const month = (est.getMonth() + 1).toString().padStart(2, '0');
  const day = est.getDate().toString().padStart(2, '0');
  const hours = est.getHours().toString().padStart(2, '0');
  const minutes = est.getMinutes().toString().padStart(2, '0');
  const seconds = est.getSeconds().toString().padStart(2, '0');
  
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds} ${is_dst ? 'EDT' : 'EST'}`;
}

function isDaylightSavingTime(date) {
  const jan = new Date(date.getFullYear(), 0, 1);
  const jul = new Date(date.getFullYear(), 6, 1);
  const std_timezone_offset = Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
  return date.getTimezoneOffset() < std_timezone_offset;
}
