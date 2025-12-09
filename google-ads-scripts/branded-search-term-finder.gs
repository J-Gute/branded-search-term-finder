const CONFIG = {
  default_days: 10,
  max_days: 365,
  require_impressions: true,
  include_pmax: true,
  include_search: true,
  fuzzy_threshold: 0.9,
  similarity_min_threshold: 0.05,
  spreadsheet_url: "URL here",
  include_summary_table: false,
  branded_terms: [
    'nx', 'n x', 'nxx', 'nx x', 'solid edge', 'solide edge', 'solidedge', 'tecnomatix', 
    'technomatix', 'simcenter', 'flotherm', 'starccm', 'star ccm', 'zona', 'teamcenter', 
    'tcx', 'zelx', 'zel x', 'opcenter', 'calibre', 'aprisa', '3dic', 'capital', 'riffyn', 
    'opscenter', 'jtopen', 'opcenterx', 'heeds', 'microred', 'valor', 'ugs', 'mendix', 
    'mendex', 'mandix', 'siemens', 'team centre', 'designcenter', 'xcelerator', 
    'digital twin', 'insightshub', 'insighthub', 'tessent', 'siemeis', 'siwmens', 
    'diemens', 'siemsn', 'soemens', 'siemn', 'siemens healthineers', 'siemens empresa', 
    'siemans', 'siemons', 'pads', 'nx cam', 'anovis', 'polarian', 'seemens', 'suemens', 
    'siem ns', 'cimens', 'siemeen', 'seimene', 'siemes', 'simens', 'siemems', 'siesmens', 
    'slemens', 'siems', 'siemense', 'ziemens', 'siemins', 'simen s', 'siemies', 'seimans', 
    'siemebs', 'siement', 'siemens', 'amesim', 'mentor', 'nastran', 'testlab', 'flomaster', 
    'xpedition', 'hyperlynx', 'simatic', 'electra', 'jt open', 'jt2go', 'insights hub', 
    'geolus', 'ugnx', 'siemense', 'di sw', 'ziemens', 'pave360', 'iray+', 'fastspice', 
    'buildingx', 'vesys', 'partquest', 'tcx', 'tcpcm', 'nx mach', 'tcvis', 'altair', 
    'dotmatics', 'culgi', 'femap', 'mindsphere', 'polarian', 'polarion', 'unigraphics', 'heeds', 'tia portal'
  ]
};

function main() {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(CONFIG.spreadsheet_url);
    const days = Math.min(Math.max(CONFIG.default_days, 1), CONFIG.max_days);
    const date_period = build_date_period(days);
    
    create_branded_terms_sheet(spreadsheet);
    
    if (CONFIG.include_pmax) {
      const pmax_query = build_pmax_query(date_period);
      run_enhanced_report(pmax_query, spreadsheet, days, 'PMAX');
    }
    
    if (CONFIG.include_search) {
      const search_query = build_search_query(date_period);
      run_enhanced_report(search_query, spreadsheet, days, 'Search');
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
    AND campaign.advertising_channel_type IN ('PERFORMANCE_MAX')`;
  
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
    segments.search_term_match_type,
    segments.keyword.info.match_type,
    segments.keyword.info.text,
    search_term_view.search_term,
    metrics.cost_micros,
    metrics.impressions,
    metrics.clicks,
    metrics.conversions,
    search_term_view.status
    FROM search_term_view
    WHERE segments.date ${date_period}`;
  
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
    
    while (report.hasNext()) {
      const row = report.next();
      row_count++;
      
      let search_term, cost, impressions, clicks, conversions, campaign_id, campaign_name, status, customer_id, customer_name;
      let search_term_match_type = '', keyword_match_type = '', keyword_text = '';
      
      if (report_type === 'PMAX') {
        search_term = row.campaignSearchTermView.searchTerm;
        cost = row.metrics.costMicros / 1000000;
        impressions = row.metrics.impressions;
        clicks = row.metrics.clicks;
        conversions = row.metrics.conversions;
        campaign_id = row.campaign.id;
        campaign_name = row.campaign.name;
        status = row.segments.searchTermTargetingStatus;
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
        status = row.searchTermView.status;
        customer_id = row.customer.id;
        customer_name = row.customer.descriptiveName;
        search_term_match_type = row.segments.searchTermMatchType || '';
        keyword_match_type = row.segments.keyword?.info?.matchType || '';
        keyword_text = row.segments.keyword?.info?.text || '';
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
        search_term, cost, impressions, clicks, conversions, status,
        search_term_match_type, keyword_match_type, keyword_text
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
        const { search_term, cost, impressions, clicks, conversions, status, search_term_match_type, keyword_match_type, keyword_text } = term_data;
        
        const cpc = clicks > 0 ? cost / clicks : 0;
        const cost_percentage = campaign_total > 0 ? (cost / campaign_total) : 0;
        const brand_analysis = analyze_brand_matching(search_term);
        const exact_match = `[${search_term}]`;
        const phrase_match = `"${search_term}"`;
        
        if (campaign_data.campaign_type === 'PMAX') {
          processed_data.push([
            campaign_data.campaign_id, campaign_data.campaign_name, search_term, status,
            brand_analysis.regex_match, brand_analysis.fuzzy_match, brand_analysis.similarity_score,
            cost, cost_percentage, impressions, clicks, cpc, conversions, exact_match, phrase_match
          ]);
        } else {
          processed_data.push([
            campaign_data.campaign_id, campaign_data.campaign_name, search_term, status,
            search_term_match_type, keyword_match_type, keyword_text,
            brand_analysis.regex_match, brand_analysis.fuzzy_match, brand_analysis.similarity_score,
            cost, cost_percentage, impressions, clicks, cpc, conversions, exact_match, phrase_match
          ]);
        }
      }
      
      let header_row = CONFIG.include_summary_table ? 15 : 11;
      if (CONFIG.include_summary_table) {
        add_summary_table(campaign_sheet, processed_data, campaign_data.campaign_type);
      }
      
      let headers = campaign_data.campaign_type === 'PMAX'
        ? ["Campaign ID", "Campaign Name", "Search Term", "Status", "Regex Match", "Fuzzy Match (more info.)", "Similarity Score", "Cost", "% of search cost", "Impressions", "Clicks", "CPC", "Conversions", "Exact Match", "Phrase Match"]
        : ["Campaign ID", "Campaign Name", "Search Term", "Status", "Search Term Match Type", "Keyword Match Type", "Keyword Text", "Regex Match", "Fuzzy Match (more info.)", "Similarity Score", "Cost", "% of search cost", "Impressions", "Clicks", "CPC", "Conversions", "Exact Match", "Phrase Match"];
      
      campaign_sheet.getRange(header_row, 1, 1, headers.length).setValues([headers]);
      
      const fuzzy_match_col = campaign_data.campaign_type === 'PMAX' ? 6 : 9;
      const fuzzy_match_cell = campaign_sheet.getRange(header_row, fuzzy_match_col);
      fuzzy_match_cell.setFormula('=HYPERLINK("https://en.wikipedia.org/wiki/Approximate_string_matching", "Fuzzy Match (more info.)")');
      
      if (processed_data.length > 0) {
        campaign_sheet.getRange(header_row + 1, 1, processed_data.length, headers.length).setValues(processed_data);
        format_results_sheet(campaign_sheet, headers.length, header_row, campaign_data.campaign_type);
      }
      
      Logger.log(`Created ${campaign_data.campaign_type} sheet for campaign ${campaign_data.campaign_id}`);
    }
    
    Logger.log(`Processed ${row_count} ${report_type} rows`);
  } catch (error) {
    Logger.log(`Error in run_enhanced_report: ${error.message}`);
    throw error;
  }
}

function add_summary_table(sheet, processed_data, campaign_type) {
  if (!CONFIG.include_summary_table) return;
  
  const branded_stats = { cost: 0, impressions: 0, clicks: 0, conversions: 0, count: 0 };
  const non_branded_stats = { cost: 0, impressions: 0, clicks: 0, conversions: 0, count: 0 };
  
  for (const row of processed_data) {
    let regex_match, fuzzy_match, cost, impressions, clicks, conversions;
    
    if (campaign_type === 'PMAX') {
      regex_match = row[4];
      fuzzy_match = row[5];
      cost = row[7];
      impressions = row[9];
      clicks = row[10];
      conversions = row[12];
    } else {
      regex_match = row[7];
      fuzzy_match = row[8];
      cost = row[10];
      impressions = row[12];
      clicks = row[13];
      conversions = row[15];
    }
    
    const is_branded = regex_match === 'likely branded' || fuzzy_match === 'likely branded' || fuzzy_match === 'possibly branded';
    
    if (is_branded) {
      branded_stats.cost += cost;
      branded_stats.impressions += impressions;
      branded_stats.clicks += clicks;
      branded_stats.conversions += conversions;
      branded_stats.count += 1;
    } else {
      non_branded_stats.cost += cost;
      non_branded_stats.impressions += impressions;
      non_branded_stats.clicks += clicks;
      non_branded_stats.conversions += conversions;
      non_branded_stats.count += 1;
    }
  }
  
  const branded_cpc = branded_stats.clicks > 0 ? branded_stats.cost / branded_stats.clicks : 0;
  const non_branded_cpc = non_branded_stats.clicks > 0 ? non_branded_stats.cost / non_branded_stats.clicks : 0;
  
  const summary_headers = ["Category", "Search Terms", "Cost", "Impressions", "Clicks", "CPC", "Conversions", "Similarity Score"];
  sheet.getRange(12, 6, 1, summary_headers.length).setValues([summary_headers]);
  
  const branded_row = ["Branded Terms", branded_stats.count, branded_stats.cost, branded_stats.impressions, branded_stats.clicks, branded_cpc, branded_stats.conversions, "-"];
  sheet.getRange(13, 6, 1, branded_row.length).setValues([branded_row]);
  
  const non_branded_row = ["Non-Branded Terms", non_branded_stats.count, non_branded_stats.cost, non_branded_stats.impressions, non_branded_stats.clicks, non_branded_cpc, non_branded_stats.conversions, "-"];
  sheet.getRange(14, 6, 1, non_branded_row.length).setValues([non_branded_row]);
  
  const summary_range = sheet.getRange(12, 6, 3, summary_headers.length);
  summary_range.setBorder(true, true, true, true, true, true);
  
  sheet.getRange(12, 6, 1, summary_headers.length).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(13, 8, 2, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(13, 11, 2, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(13, 7, 2, 1).setNumberFormat('#,##0');
  sheet.getRange(13, 9, 2, 1).setNumberFormat('#,##0');
  sheet.getRange(13, 10, 2, 1).setNumberFormat('#,##0');
  sheet.getRange(13, 12, 2, 1).setNumberFormat('0.00');
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

function analyze_brand_matching(search_term) {
  if (!search_term || typeof search_term !== 'string') {
    return { regex_match: '', fuzzy_match: '', similarity_score: 0 };
  }
  
  if (is_entirely_non_latin(search_term)) {
    return { regex_match: '', fuzzy_match: 'not evaluated', similarity_score: 0 };
  }
  
  const regex_result = check_regex_brand_match(search_term);
  if (regex_result === 'likely branded') {
    return { regex_match: regex_result, fuzzy_match: 'likely branded', similarity_score: 1.0 };
  }
  
  const similarity_score = BRANDED_FUZZY_3(search_term, CONFIG.branded_terms);
  const fuzzy_classification = classify_fuzzy_result(similarity_score);
  
  return {
    regex_match: regex_result,
    fuzzy_match: fuzzy_classification,
    similarity_score: Math.round(similarity_score * 1000) / 1000
  };
}

function is_entirely_non_latin(search_term) {
  const cleaned_term = search_term.replace(/[\s\-_.]/g, '');
  if (cleaned_term.length === 0) return false;
  const non_latin_regex = /^[\u0400-\u04FF\u0500-\u052F\u2DE0-\u2DFF\uA640-\uA69F\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\u0600-\u06FF]+$/;
  return non_latin_regex.test(cleaned_term);
}

function check_regex_brand_match(search_term) {
  const search_lower = search_term.toLowerCase();
  
  for (const branded_term of CONFIG.branded_terms) {
    const branded_lower = branded_term.toLowerCase();
    
    if (search_lower === branded_lower) return 'likely branded';
    if (branded_lower.length > 3 && search_lower.includes(branded_lower)) return 'likely branded';
    
    if (branded_lower.length <= 3) {
      const word_boundary_regex = new RegExp(`\\b${escape_regex(branded_lower)}\\b`, 'i');
      if (word_boundary_regex.test(search_term)) return 'likely branded';
    }
  }
  
  return '';
}

function classify_fuzzy_result(similarity_score) {
  if (similarity_score >= CONFIG.fuzzy_threshold) return 'possibly branded';
  if (similarity_score >= CONFIG.similarity_min_threshold) return 'unlikely';
  return '';
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
  let column_widths;
  
  if (campaign_type === 'PMAX') {
    column_widths = { 1: 110, 2: 210, 3: 260, 4: 130, 5: 110, 6: 160, 7: 110, 8: 90, 9: 130, 10: 90, 11: 70, 12: 80, 13: 90, 14: 150, 15: 150 };
  } else {
    column_widths = { 1: 110, 2: 210, 3: 260, 4: 130, 5: 150, 6: 150, 7: 200, 8: 110, 9: 160, 10: 110, 11: 90, 12: 130, 13: 90, 14: 70, 15: 80, 16: 90, 17: 150, 18: 150 };
  }
  
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
    
    if (campaign_type === 'PMAX') {
      sheet.getRange(header_row + 1, 8, data_rows, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(header_row + 1, 9, data_rows, 1).setNumberFormat('0.00%');
      sheet.getRange(header_row + 1, 12, data_rows, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(header_row + 1, 7, data_rows, 1).setNumberFormat('0.000');
      sheet.getRange(header_row + 1, 10, data_rows, 1).setNumberFormat('#,##0');
      sheet.getRange(header_row + 1, 11, data_rows, 1).setNumberFormat('#,##0');
      sheet.getRange(header_row + 1, 13, data_rows, 1).setNumberFormat('0.00');
    } else {
      sheet.getRange(header_row + 1, 11, data_rows, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(header_row + 1, 12, data_rows, 1).setNumberFormat('0.00%');
      sheet.getRange(header_row + 1, 15, data_rows, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(header_row + 1, 10, data_rows, 1).setNumberFormat('0.000');
      sheet.getRange(header_row + 1, 13, data_rows, 1).setNumberFormat('#,##0');
      sheet.getRange(header_row + 1, 14, data_rows, 1).setNumberFormat('#,##0');
      sheet.getRange(header_row + 1, 16, data_rows, 1).setNumberFormat('0.00');
    }
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
