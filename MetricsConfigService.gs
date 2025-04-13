/**
 * MetricsConfigService.gs
 * Service for configuring metrics email reports and data exports
 * 
 * This service provides functions for:
 * - Creating and sending email reports with metrics
 * - Configuring export formats and scheduling
 * - Building email templates with metrics data
 * - Managing export configurations
 * 
 * Part of the 988 Crisis Line Team Lead Dashboard
 * 
 * @version 1.0.0
 * @author Crisis Line Team Lead Dashboard System
 */

/**
 * Default configuration for email reports and exports
 */
const DEFAULT_CONFIG = {
  email: {
    reportName: "988 Crisis Line Metrics Report",
    emailSubject: "988 Crisis Line Weekly Metrics Report",
    reportType: "comprehensive", // comprehensive, team, individual
    reportFormat: "rich", // rich, simple, text
    dateRange: "pastWeek", // pastWeek, pastMonth, pastQuarter, pastYear, custom
    metrics: ["Answer Rate", "Talk Time", "After Call Work", "On Queue Time", "Calls Handled"],
    charts: ["Call Volume", "Answer Rate"],
    schedule: {
      enabled: false,
      type: "weekly", // weekly, monthly, custom
      days: [1], // 0=Sunday, 1=Monday, etc.
      time: "08:00", // 24-hour format
      frequency: "weekly" // weekly, monthly
    },
    recipients: [],
    includeMyself: true,
    replyToEmail: "",
    fromName: "988 Crisis Line Analytics",
    attachmentType: "none", // none, pdf, spreadsheet, both
    introMessage: "Here's your weekly metrics report for the 988 Crisis Line team."
  },
  export: {
    exportName: "988 Crisis Line Metrics",
    exportType: "team", // team, individual, raw, comprehensive
    dateRange: "pastWeek",
    groupBy: "day",
    fields: ["Answer Rate", "Talk Time", "After Call Work", "On Queue Time", "Calls Offered", "Calls Accepted", "Staff Count"],
    format: "spreadsheet", // spreadsheet, csv, pdf
    includeCharts: true,
    includeRawData: true,
    includeComparisons: true,
    scheduleExport: false
  }
};

/**
 * Loads email configuration from user properties
 * @return {Object} Email configuration
 */
function loadEmailConfig() {
  try {
    // Get user-specific configuration
    const userEmail = UserService.getCurrentUserEmail();
    
    // Try to get stored config or use default
    let config = StorageService.getProperty(`${userEmail}_metrics_email_config`);
    if (!config) {
      return DEFAULT_CONFIG.email;
    }
    
    return JSON.parse(config);
  } catch (error) {
    console.error("Error loading email configuration: " + error.message);
    return DEFAULT_CONFIG.email;
  }
}

/**
 * Saves email configuration to user properties
 * @param {Object} config - Email configuration to save
 * @return {Object} Result with success flag and message
 */
function saveEmailConfiguration(config) {
  try {
    // Validate config
    validateEmailConfig(config);
    
    // Get user-specific configuration
    const userEmail = UserService.getCurrentUserEmail();
    
    // Save to user's properties
    StorageService.setProperty(
      `${userEmail}_metrics_email_config`, 
      JSON.stringify(config)
    );
    
    // Set up email schedule trigger if enabled
    if (config.schedule && config.schedule.enabled) {
      setupEmailScheduleTrigger(config.schedule);
    } else {
      removeEmailScheduleTrigger();
    }
    
    return {
      success: true,
      message: "Email configuration saved successfully"
    };
  } catch (error) {
    console.error("Error saving email configuration: " + error.message);
    
    return {
      success: false,
      message: "Error: " + error.message
    };
  }
}

/**
 * Validates email configuration
 * @param {Object} config - Email configuration to validate
 * @throws {Error} If configuration is invalid
 */
function validateEmailConfig(config) {
  // Check required fields
  if (!config.reportName) throw new Error("Report name is required");
  if (!config.emailSubject) throw new Error("Email subject is required");
  if (!config.reportType) throw new Error("Report type is required");
  if (!config.reportFormat) throw new Error("Report format is required");
  if (!config.dateRange) throw new Error("Date range is required");
  
  // Validate recipients if provided
  if (config.recipients && Array.isArray(config.recipients)) {
    config.recipients.forEach(recipient => {
      if (!recipient.email) throw new Error("All recipients must have an email address");
      if (!isValidEmail(recipient.email)) throw new Error(`Invalid email address: ${recipient.email}`);
    });
  }
  
  // Validate schedule if enabled
  if (config.schedule && config.schedule.enabled) {
    if (!config.schedule.type) throw new Error("Schedule type is required");
    
    if (config.schedule.type === "weekly" && (!config.schedule.days || !config.schedule.days.length)) {
      throw new Error("At least one day must be selected for weekly schedule");
    }
    
    if (config.schedule.type === "monthly" && !config.schedule.day) {
      throw new Error("Day of month is required for monthly schedule");
    }
    
    if (config.schedule.type === "custom") {
      if (!config.schedule.frequency) throw new Error("Frequency is required for custom schedule");
      if (config.schedule.frequency === "weekly" && (!config.schedule.days || !config.schedule.days.length)) {
        throw new Error("At least one day must be selected for weekly custom schedule");
      }
      if (config.schedule.frequency === "monthly" && !config.schedule.day) {
        throw new Error("Day of month is required for monthly custom schedule");
      }
    }
  }
  
  // Validate at least one metric is selected
  if (!config.metrics || !Array.isArray(config.metrics) || config.metrics.length === 0) {
    throw new Error("At least one metric must be selected");
  }
}

/**
 * Check if an email address is valid
 * @param {string} email - Email address to validate
 * @return {boolean} Whether the email is valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Set up email schedule trigger
 * @param {Object} scheduleConfig - Schedule configuration
 */
function setupEmailScheduleTrigger(scheduleConfig) {
  try {
    // Remove any existing triggers
    removeEmailScheduleTrigger();
    
    // Set up new trigger based on schedule type
    switch (scheduleConfig.type) {
      case "weekly":
        setupWeeklyTrigger(scheduleConfig);
        break;
      case "monthly":
        setupMonthlyTrigger(scheduleConfig);
        break;
      case "custom":
        setupCustomTrigger(scheduleConfig);
        break;
      default:
        throw new Error("Invalid schedule type: " + scheduleConfig.type);
    }
  } catch (error) {
    console.error("Error setting up email schedule trigger: " + error.message);
    throw error;
  }
}

/**
 * Set up weekly email trigger
 * @param {Object} config - Schedule configuration
 */
function setupWeeklyTrigger(config) {
  // Get time components
  const timeComponents = parseTimeString(config.time || "08:00");
  
  // Create a trigger for each selected day
  config.days.forEach(day => {
    ScriptApp.newTrigger('sendScheduledEmailReport')
      .timeBased()
      .onWeekDay(getWeekdayFromNumber(day))
      .atHour(timeComponents.hour)
      .nearMinute(timeComponents.minute)
      .create();
  });
}

/**
 * Set up monthly email trigger
 * @param {Object} config - Schedule configuration
 */
function setupMonthlyTrigger(config) {
  // Monthly triggers are handled by a daily trigger that checks the day
  ScriptApp.newTrigger('checkAndSendMonthlyEmailReport')
    .timeBased()
    .everyDays(1)
    .atHour(1) // Run check at 1 AM
    .create();
}

/**
 * Set up custom schedule trigger
 * @param {Object} config - Schedule configuration
 */
function setupCustomTrigger(config) {
  if (config.frequency === "weekly") {
    setupWeeklyTrigger(config);
  } else if (config.frequency === "monthly") {
    setupMonthlyTrigger(config);
  }
}

/**
 * Remove existing email schedule triggers
 */
function removeEmailScheduleTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendScheduledEmailReport' || 
        trigger.getHandlerFunction() === 'checkAndSendMonthlyEmailReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Parse time string in format "HH:MM"
 * @param {string} timeString - Time string
 * @return {Object} Object with hour and minute components
 */
function parseTimeString(timeString) {
  try {
    const parts = timeString.split(":");
    return {
      hour: parseInt(parts[0], 10),
      minute: parseInt(parts[1], 10)
    };
  } catch (error) {
    console.error("Error parsing time string: " + error.message);
    // Default to 8:00 AM
    return { hour: 8, minute: 0 };
  }
}

/**
 * Get ScriptApp weekday enum from number (0-6, Sunday-Saturday)
 * @param {number} dayNumber - Day number (0-6)
 * @return {Weekday} ScriptApp weekday enum
 */
function getWeekdayFromNumber(dayNumber) {
  const weekdays = [
    ScriptApp.WeekDay.SUNDAY,
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY,
    ScriptApp.WeekDay.SATURDAY
  ];
  
  return weekdays[dayNumber] || ScriptApp.WeekDay.MONDAY; // Default to Monday
}

/**
 * Get start date based on date range setting
 * @param {string} dateRange - Date range setting
 * @param {Date} endDate - End date
 * @return {Date} Start date
 */
function getStartDateFromConfig(dateRange, endDate) {
  const startDate = new Date(endDate);
  
  switch (dateRange) {
    case "pastWeek":
      startDate.setDate(startDate.getDate() - 7);
      break;
    case "pastMonth":
      startDate.setDate(startDate.getDate() - 30);
      break;
    case "pastQuarter":
      startDate.setDate(startDate.getDate() - 90);
      break;
    case "pastYear":
      startDate.setDate(startDate.getDate() - 365);
      break;
    case "custom":
      // Custom date range should be handled separately
      break;
    default:
      startDate.setDate(startDate.getDate() - 7); // Default to past week
  }
  
  // Set time to start of day
  startDate.setHours(0, 0, 0, 0);
  
  return startDate;
}

/**
 * Format date for display
 * @param {Date} date - Date to format
 * @return {string} Formatted date
 */
function formatDateForDisplay(date) {
  try {
    return date.toLocaleDateString();
  } catch (error) {
    console.error("Error formatting date: " + error.message);
    return date.toString();
  }
}

/**
 * Format email subject with date range
 * @param {string} subject - Email subject template
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} Formatted email subject
 */
function formatEmailSubject(subject, startDate, endDate) {
  try {
    // Format dates for display
    const startStr = formatDateForDisplay(startDate);
    const endStr = formatDateForDisplay(endDate);
    
    // Replace placeholders if they exist
    if (subject.includes("{dateRange}")) {
      return subject.replace("{dateRange}", `${startStr} - ${endStr}`);
    } else if (subject.includes("{start}") && subject.includes("{end}")) {
      return subject.replace("{start}", startStr).replace("{end}", endStr);
    } else {
      // Append date range if no placeholders
      return `${subject}: ${startStr} - ${endStr}`;
    }
  } catch (error) {
    console.error("Error formatting email subject: " + error.message);
    return subject;
  }
}

/**
 * Generate email content based on configuration
 * @param {Object} config - Email configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} HTML email content
 */
function generateEmailContent(config, startDate, endDate) {
  try {
    // Generate HTML or text content based on report format
    if (config.reportFormat === "rich") {
      return generateRichEmailContent(config, startDate, endDate);
    } else if (config.reportFormat === "simple") {
      return generateSimpleEmailContent(config, startDate, endDate);
    } else {
      return generateTextEmailContent(config, startDate, endDate);
    }
  } catch (error) {
    console.error("Error generating email content: " + error.message);
    throw error;
  }
}

/**
 * Generate rich HTML email content
 * @param {Object} config - Email configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} Rich HTML email content
 */
function generateRichEmailContent(config, startDate, endDate) {
  try {
    // Get metrics data for the date range
    const metricsData = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Get team members data if needed
    let teamMembersData = null;
    if (config.reportType === 'individual' || config.reportType === 'comprehensive') {
      teamMembersData = TeamService.getTeamMembersWithMetrics(startDate, endDate);
    }
    
    // Build rich HTML email
    return buildRichHtmlEmail(config, metricsData, teamMembersData, startDate, endDate);
  } catch (error) {
    console.error("Error generating rich email content: " + error.message);
    throw error;
  }
}

/**
 * Generate simple HTML email content
 * @param {Object} config - Email configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} Simple HTML email content
 */
function generateSimpleEmailContent(config, startDate, endDate) {
  try {
    // Get metrics data for the date range
    const metricsData = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Get team members data if needed
    let teamMembersData = null;
    if (config.reportType === 'individual' || config.reportType === 'comprehensive') {
      teamMembersData = TeamService.getTeamMembersWithMetrics(startDate, endDate);
    }
    
    // Build simple HTML email
    return buildSimpleHtmlEmail(config, metricsData, teamMembersData, startDate, endDate);
  } catch (error) {
    console.error("Error generating simple email content: " + error.message);
    throw error;
  }
}

/**
 * Build rich HTML email with charts and advanced formatting
 * @param {Object} config - Email configuration
 * @param {Object} metricsData - Team metrics data
 * @param {Array} teamMembersData - Team members data (optional)
 * @param {Date} startDate - Start date for the report
 * @param {Date} endDate - End date for the report
 * @return {string} HTML email content
 */
function buildRichHtmlEmail(config, metricsData, teamMembersData, startDate, endDate) {
  // Format dates for display
  const dateRange = `${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}`;
  const generationDate = formatDateForDisplay(new Date());
  const generationTime = new Date().toLocaleTimeString();
  
  // Start building HTML
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${config.reportName}: ${dateRange}</title>
      <style>
        body {
          font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
          margin: 0;
          padding: 0;
          background-color: #f8f9fa;
          color: #202124;
          -webkit-font-smoothing: antialiased;
        }
        
        .container {
          max-width: 600px;
          margin: 0 auto;
          background-color: #ffffff;
        }
        
        .email-header {
          background-color: #4285F4;
          padding: 20px;
          text-align: center;
          color: white;
        }
        
        .email-content {
          padding: 20px;
        }
        
        .email-title {
          font-size: 20px;
          font-weight: bold;
          margin-bottom: 15px;
          color: #202124;
        }
        
        .email-section {
          margin-bottom: 30px;
        }
        
        .email-section-title {
          font-size: 16px;
          font-weight: bold;
          color: #4285F4;
          margin-bottom: 15px;
          padding-bottom: 5px;
          border-bottom: 1px solid #dadce0;
        }
        
        .email-metrics {
          display: flex;
          flex-wrap: wrap;
          margin: 0 -10px;
        }
        
        .email-metric {
          width: calc(50% - 20px);
          margin: 0 10px 20px;
          text-align: center;
          padding: 15px;
          background-color: #f1f3f4;
          border-radius: 4px;
        }
        
        .email-metric-value {
          font-size: 24px;
          font-weight: bold;
          margin-bottom: 5px;
          color: #4285F4;
        }
        
        .email-metric-label {
          font-size: 12px;
          color: #5F6368;
        }
        
        .email-chart {
          margin: 20px 0;
          padding: 15px;
          background-color: #f1f3f4;
          text-align: center;
          border-radius: 4px;
        }
        
        .email-table {
          width: 100%;
          border-collapse: collapse;
          margin: 15px 0;
        }
        
        .email-table th {
          background-color: #f1f3f4;
          padding: 10px;
          text-align: left;
          font-size: 12px;
          color: #5F6368;
        }
        
        .email-table td {
          padding: 10px;
          border-top: 1px solid #dadce0;
          font-size: 13px;
        }
        
        .email-footer {
          padding: 20px;
          text-align: center;
          font-size: 12px;
          color: #5F6368;
          background-color: #f8f9fa;
        }
        
        /* Responsive adjustments */
        @media only screen and (max-width: 600px) {
          .email-metric {
            width: 100%;
            margin-right: 0;
            margin-left: 0;
          }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <!-- Email Header -->
        <div class="email-header">
          <div style="font-size: 20px; font-weight: bold;">${config.reportName}</div>
        </div>
        
        <!-- Email Content -->
        <div class="email-content">
          <div class="email-title">${config.reportName}: ${dateRange}</div>
          
          <p style="margin-bottom: 20px;">${config.introMessage || 'Here\'s the latest metrics report for your team.'}</p>
          
          <!-- Key Metrics Section -->
          <div class="email-section">
            <div class="email-section-title">Key Performance Indicators</div>
            
            <div class="email-metrics">
  `;
  
  // Add metrics
  config.metrics.forEach(metric => {
    let value = '—';
    let color = '#4285F4'; // Default color
    
    switch (metric) {
      case 'Answer Rate':
        value = (metricsData.answerRate * 100).toFixed(1) + '%';
        // Color based on value
        if (metricsData.answerRate >= 0.95) {
          color = '#34A853'; // Good
        } else if (metricsData.answerRate >= 0.9) {
          color = '#FBBC05'; // Warning
        } else {
          color = '#EA4335'; // Poor
        }
        break;
      case 'Talk Time':
        value = metricsData.averageTalkTime.toFixed(1) + ' min';
        // Color based on value (15-20 min is ideal)
        if (metricsData.averageTalkTime >= 15 && metricsData.averageTalkTime <= 20) {
          color = '#34A853'; // Good
        } else if ((metricsData.averageTalkTime >= 13 && metricsData.averageTalkTime < 15) || 
                  (metricsData.averageTalkTime > 20 && metricsData.averageTalkTime <= 22)) {
          color = '#FBBC05'; // Warning
        } else {
          color = '#EA4335'; // Poor
        }
        break;
      case 'After Call Work':
        value = metricsData.acwAverage.toFixed(1) + ' min';
        // Color based on value (lower is better)
        if (metricsData.acwAverage <= 5) {
          color = '#34A853'; // Good
        } else if (metricsData.acwAverage <= 7) {
          color = '#FBBC05'; // Warning
        } else {
          color = '#EA4335'; // Poor
        }
        break;
      case 'On Queue Time':
        value = (metricsData.onQueuePercentage * 100).toFixed(1) + '%';
        // Color based on value (higher is better)
        if (metricsData.onQueuePercentage >= 0.65) {
          color = '#34A853'; // Good
        } else if (metricsData.onQueuePercentage >= 0.6) {
          color = '#FBBC05'; // Warning
        } else {
          color = '#EA4335'; // Poor
        }
        break;
      case 'Calls Handled':
        value = metricsData.callsAccepted.toLocaleString();
        break;
      case 'Calls Offered':
        value = metricsData.callsOffered.toLocaleString();
        break;
      case 'Interacting Time':
        value = (metricsData.interactingTimePercentage * 100).toFixed(1) + '%';
        break;
      // Add other metrics as needed...
    }
    
    html += `
      <div class="email-metric">
        <div class="email-metric-value" style="color: ${color};">${value}</div>
        <div class="email-metric-label">${metric}</div>
      </div>
    `;
  });
  
  html += `
            </div>
  `;
  
  // Add charts if selected
  if (config.charts && config.charts.includes('Call Volume')) {
    html += `
      <div class="email-chart">
        <div style="font-weight: bold; margin-bottom: 10px;">Call Volume Trend</div>
        <img src="https://chart.googleapis.com/chart?cht=lc&chs=500x200&chd=t:${generateChartData(metricsData, 'callsOffered')}&chxt=x,y&chxl=0:|${generateChartLabels(metricsData)}&chco=4285F4&chf=bg,s,f8f9fa&chds=a" width="100%" alt="Call Volume Chart">
      </div>
    `;
  }
  
  if (config.charts && config.charts.includes('Answer Rate')) {
    html += `
      <div class="email-chart">
        <div style="font-weight: bold; margin-bottom: 10px;">Answer Rate Trend</div>
        <img src="https://chart.googleapis.com/chart?cht=lc&chs=500x200&chd=t:${generateChartData(metricsData, 'answerRate', 100)}&chxt=x,y&chxl=0:|${generateChartLabels(metricsData)}&chco=34A853&chf=bg,s,f8f9fa&chds=80,100&chm=h,FBBC05,0,0.9,1,1:dash&chm=h,EA4335,0,0.85,1,1:dash" width="100%" alt="Answer Rate Chart">
      </div>
    `;
  }
  
  html += `
          </div>
  `;
  
  // Add team performance section if applicable
  if ((config.reportType === 'comprehensive' || config.reportType === 'individual') && 
      teamMembersData && teamMembersData.length > 0) {
    
    html += `
      <!-- Team Performance Section -->
      <div class="email-section">
        <div class="email-section-title">Team Performance</div>
        
        <table class="email-table">
          <thead>
            <tr>
              <th style="background-color: #f1f3f4; padding: 10px; text-align: left; font-size: 12px; color: #5F6368;">Team Member</th>
              <th style="background-color: #f1f3f4; padding: 10px; text-align: left; font-size: 12px; color: #5F6368;">Calls Handled</th>
              <th style="background-color: #f1f3f4; padding: 10px; text-align: left; font-size: 12px; color: #5F6368;">Answer Rate</th>
                <th style="background-color: #f1f3f4; padding: 10px; text-align: left; font-size: 12px; color: #5F6368;">Talk Time</th>
                <th style="background-color: #f1f3f4; padding: 10px; text-align: left; font-size: 12px; color: #5F6368;">Score</th>
              </tr>
            </thead>
            <tbody>
      `;
      
      // Add team member rows - limit to top 5 performers
      teamMembersData
        .sort((a, b) => b.score - a.score)
        .slice(0, 5)
        .forEach((member, index) => {
          html += `
            <tr>
              <td style="padding: 10px; border-top: 1px solid #dadce0; font-size: 13px;">${member.name}</td>
              <td style="padding: 10px; border-top: 1px solid #dadce0; font-size: 13px;">${member.callsAccepted}</td>
              <td style="padding: 10px; border-top: 1px solid #dadce0; font-size: 13px;">${(member.answerRate * 100).toFixed(1)}%</td>
              <td style="padding: 10px; border-top: 1px solid #dadce0; font-size: 13px;">${member.averageTalkTime.toFixed(1)} min</td>
              <td style="padding: 10px; border-top: 1px solid #dadce0; font-size: 13px;">${Math.round(member.score)}</td>
            </tr>
          `;
        });
      
      html += `
            </tbody>
          </table>
        </div>
      `;
    }
  
  // Add footer
  html += `
      </div>
      
      <!-- Email Footer -->
      <div class="email-footer">
        <p>This is an automated report from the 988 Crisis Line Metrics System.</p>
        <p style="margin-top: 10px;">Generated on ${generationDate} at ${generationTime}</p>
      </div>
    </div>
  `;
  
  return html;
}

/**
 * Build simple HTML email with tables but no charts
 * @param {Object} config - Email configuration
 * @param {Object} metricsData - Team metrics data
 * @param {Array} teamMembersData - Team members data (optional)
 * @param {Date} startDate - Start date for the report
 * @param {Date} endDate - End date for the report
 * @return {string} HTML email content
 */
function buildSimpleHtmlEmail(config, metricsData, teamMembersData, startDate, endDate) {
  // Format dates for display
  const dateRange = `${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}`;
  const generationDate = formatDateForDisplay(new Date());
  const generationTime = new Date().toLocaleTimeString();
  
  // Start building HTML
  let html = `
    <div style="font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif; max-width: 600px; margin: 0 auto;">
      <!-- Email Header -->
      <div style="background-color: #4285F4; padding: 15px; text-align: center;">
        <div style="color: white; font-size: 20px; font-weight: bold;">${config.reportName}</div>
      </div>
      
      <!-- Email Content -->
      <div style="padding: 20px; background-color: #f8f9fa;">
        <div style="font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #202124;">${config.reportName}: ${dateRange}</div>
        
        <p style="margin-bottom: 20px; color: #202124;">${config.introMessage || 'Here\'s the latest metrics report for your team.'}</p>
        
        <!-- Key Metrics Table -->
        <div style="background-color: white; border-radius: 4px; padding: 15px; margin-bottom: 20px; box-shadow: 0 1px 2px rgba(60, 64, 67, 0.3);">
          <div style="font-size: 16px; font-weight: bold; margin-bottom: 10px; color: #4285F4;">Key Performance Indicators</div>
          
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 10px;">
            <thead>
              <tr>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: left; font-size: 12px; color: #5F6368;">Metric</th>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: right; font-size: 12px; color: #5F6368;">Value</th>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: right; font-size: 12px; color: #5F6368;">Goal</th>
              </tr>
            </thead>
            <tbody>
  `;
  
  // Add metrics rows
  const metricGoals = {
    'Answer Rate': '≥95%',
    'Talk Time': '15-20 min',
    'After Call Work': '≤5 min',
    'On Queue Time': '≥65%',
    'Calls Handled': 'N/A',
    'Calls Offered': 'N/A',
    'Interacting Time': '≥50%'
  };
  
  config.metrics.forEach(metric => {
    let value = '—';
    
    switch (metric) {
      case 'Answer Rate':
        value = (metricsData.answerRate * 100).toFixed(1) + '%';
        break;
      case 'Talk Time':
        value = metricsData.averageTalkTime.toFixed(1) + ' min';
        break;
      case 'After Call Work':
        value = metricsData.acwAverage.toFixed(1) + ' min';
        break;
      case 'On Queue Time':
        value = (metricsData.onQueuePercentage * 100).toFixed(1) + '%';
        break;
      case 'Calls Handled':
        value = metricsData.callsAccepted.toLocaleString();
        break;
      case 'Calls Offered':
        value = metricsData.callsOffered.toLocaleString();
        break;
      case 'Interacting Time':
        value = (metricsData.interactingTimePercentage * 100).toFixed(1) + '%';
        break;
      // Add other metrics as needed...
    }
    
    html += `
      <tr>
        <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px;">${metric}</td>
        <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px; text-align: right; font-weight: bold;">${value}</td>
        <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px; text-align: right; color: #5F6368;">${metricGoals[metric] || 'N/A'}</td>
      </tr>
    `;
  });
  
  html += `
            </tbody>
          </table>
        </div>
  `;
  
  // Add team performance table if applicable
  if (config.reportType === 'comprehensive' || config.reportType === 'individual') {
    if (teamMembersData && teamMembersData.length > 0) {
      html += `
        <!-- Team Performance Section -->
        <div style="background-color: white; border-radius: 4px; padding: 15px; margin-bottom: 20px; box-shadow: 0 1px 2px rgba(60, 64, 67, 0.3);">
          <div style="font-size: 16px; font-weight: bold; margin-bottom: 10px; color: #4285F4;">Team Performance</div>
          
          <table style="width: 100%; border-collapse: collapse;">
            <thead>
              <tr>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: left; font-size: 12px; color: #5F6368;">Team Member</th>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: right; font-size: 12px; color: #5F6368;">Calls</th>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: right; font-size: 12px; color: #5F6368;">Answer Rate</th>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: right; font-size: 12px; color: #5F6368;">Talk Time</th>
                <th style="background-color: #f1f3f4; padding: 8px; text-align: right; font-size: 12px; color: #5F6368;">Score</th>
              </tr>
            </thead>
            <tbody>
      `;
      
      // Add team member rows - show all team members in simple format
      teamMembersData
        .sort((a, b) => b.score - a.score)
        .forEach((member) => {
          html += `
            <tr>
              <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px;">${member.name}</td>
              <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px; text-align: right;">${member.callsAccepted}</td>
              <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px; text-align: right;">${(member.answerRate * 100).toFixed(1)}%</td>
              <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px; text-align: right;">${member.averageTalkTime.toFixed(1)} min</td>
              <td style="padding: 8px; border-top: 1px solid #dadce0; font-size: 13px; text-align: right; font-weight: bold;">${Math.round(member.score)}</td>
            </tr>
          `;
        });
      
      html += `
            </tbody>
          </table>
        </div>
      `;
    }
  }
  
  // Add footer
  html += `
      </div>
      
      <!-- Email Footer -->
      <div style="padding: 15px; text-align: center; font-size: 11px; color: #5F6368;">
        <p>This is an automated report from the 988 Crisis Line Metrics System.</p>
        <p style="margin-top: 8px;">Generated on ${generationDate} at ${generationTime}</p>
      </div>
    </div>
  `;
  
  return html;
}

/**
 * Generate plain text email content
 * @param {Object} config - Email configuration
 * @param {Date} startDate - Start date for the report
 * @param {Date} endDate - End date for the report
 * @return {string} Plain text email content
 */
function generateTextEmailContent(config, startDate, endDate) {
  try {
    // Get metrics data for the date range
    const metricsData = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Get team members data if needed
    let teamMembersData = null;
    if (config.reportType === 'individual' || config.reportType === 'comprehensive') {
      teamMembersData = TeamService.getTeamMembersWithMetrics(startDate, endDate);
    }
    
    // Format dates for display
    const dateRange = `${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}`;
    
    // Build text content
    let text = `${config.reportName}: ${dateRange}\n`;
    text += `=`.repeat(text.length) + '\n\n';
    
    text += `${config.introMessage || 'Here\'s the latest metrics report for your team.'}\n\n`;
    
    // Key metrics section
    text += 'KEY PERFORMANCE INDICATORS\n';
    text += '-------------------------\n';
    
    config.metrics.forEach(metric => {
      let value = '—';
      
      switch (metric) {
        case 'Answer Rate':
          value = (metricsData.answerRate * 100).toFixed(1) + '%';
          break;
        case 'Talk Time':
          value = metricsData.averageTalkTime.toFixed(1) + ' min';
          break;
        case 'After Call Work':
          value = metricsData.acwAverage.toFixed(1) + ' min';
          break;
        case 'On Queue Time':
          value = (metricsData.onQueuePercentage * 100).toFixed(1) + '%';
          break;
        case 'Calls Handled':
          value = metricsData.callsAccepted.toLocaleString();
          break;
        case 'Calls Offered':
          value = metricsData.callsOffered.toLocaleString();
          break;
        // Add other metrics as needed...
      }
      
      text += `${metric}: ${value}\n`;
    });
    
    text += '\n';
    
    // Team performance section if applicable
    if ((config.reportType === 'comprehensive' || config.reportType === 'individual') && 
        teamMembersData && teamMembersData.length > 0) {
      
      text += 'TEAM PERFORMANCE\n';
      text += '----------------\n';
      
      // Column headers
      text += 'Team Member'.padEnd(25) + 
              'Calls'.padStart(10) + 
              'Answer Rate'.padStart(15) + 
              'Talk Time'.padStart(15) + 
              'Score'.padStart(10) + '\n';
      
      text += '-'.repeat(75) + '\n';
      
      // Add team member rows - top 5 performers for text emails
      teamMembersData
        .sort((a, b) => b.score - a.score)
        .slice(0, 5)
        .forEach(member => {
          text += member.name.padEnd(25) + 
                  member.callsAccepted.toString().padStart(10) +
                  (member.answerRate * 100).toFixed(1).padStart(12) + '%' +
                  member.averageTalkTime.toFixed(1).padStart(12) + ' min' +
                  Math.round(member.score).toString().padStart(10) + '\n';
        });
    }
    
    // Footer
    text += '\n\n';
    text += 'This is an automated report from the 988 Crisis Line Metrics System.\n';
    text += `Generated on ${formatDateForDisplay(new Date())} at ${new Date().toLocaleTimeString()}`;
    
    return text;
  } catch (error) {
    console.error("Error generating text email content: " + error.message);
    throw error;
  }
}

/**
 * Send scheduled email report - called by time-based trigger
 */
function sendScheduledEmailReport() {
  try {
    // Load email config
    const config = loadEmailConfig();
    
    // Skip if scheduling is disabled
    if (!config.schedule || !config.schedule.enabled) {
      console.log("Email scheduling is disabled.");
      return;
    }
    
    // Get date range
    const endDate = new Date();
    const startDate = getStartDateFromConfig(config.dateRange, endDate);
    
    // Generate email content
    const htmlContent = generateEmailContent(config, startDate, endDate);
    const textContent = generateTextEmailContent(config, startDate, endDate);
    
    // Format the subject with date range
    const subject = formatEmailSubject(config.emailSubject, startDate, endDate);
    
    // Get recipients list
    const recipients = getRecipientsListForSending(config);
    
    // If no recipients, log and exit
    if (recipients.length === 0) {
      console.log("No recipients configured for the email report.");
      return;
    }
    
    // Create attachment(s) if needed
    const attachments = [];
    if (config.attachmentType !== 'none') {
      if (config.attachmentType === 'pdf' || config.attachmentType === 'both') {
        const pdfBlob = generatePdfReport(config, startDate, endDate);
        if (pdfBlob) {
          attachments.push(pdfBlob);
        }
      }
      
      if (config.attachmentType === 'spreadsheet' || config.attachmentType === 'both') {
        const spreadsheetBlob = generateSpreadsheetReport(config, startDate, endDate);
        if (spreadsheetBlob) {
          attachments.push(spreadsheetBlob);
        }
      }
    }
    
    // Send email
    const fromName = config.fromName || "988 Crisis Line Analytics";
    
    if (config.reportFormat === 'rich') {
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: subject,
        htmlBody: htmlContent,
        name: fromName,
        replyTo: config.replyToEmail || '',
        attachments: attachments
      });
    } else {
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: subject,
        body: textContent,
        name: fromName,
        replyTo: config.replyToEmail || '',
        attachments: attachments
      });
    }
    
    console.log(`Email report sent to ${recipients.length} recipients.`);
  } catch (error) {
    console.error("Error sending scheduled email report: " + error.message);
    
    // Send error notification to admin or current user
    try {
      const userEmail = UserService.getCurrentUserEmail();
      
      MailApp.sendEmail({
        to: userEmail,
        subject: "Error: 988 Crisis Line Metrics Report Failed",
        body: `There was an error sending the scheduled metrics report:\n\n${error.message}\n\nPlease check the configuration.`,
        name: "988 Crisis Line Analytics System"
      });
    } catch (e) {
      console.error("Failed to send error notification: " + e.message);
    }
  }
}

/**
 * Check if today is the scheduled day for monthly email and send if it is
 * Called by a daily trigger
 */
function checkAndSendMonthlyEmailReport() {
  try {
    // Load config
    const config = loadEmailConfig();
    
    // Skip if scheduling is disabled
    if (!config.schedule || !config.schedule.enabled) {
      return;
    }
    
    // Skip if not a monthly schedule
    if (config.schedule.type !== 'monthly' && 
        !(config.schedule.type === 'custom' && config.schedule.frequency === 'monthly')) {
      return;
    }
    
    // Get current day of month
    const today = new Date();
    const dayOfMonth = today.getDate();
    
    // Get scheduled day
    let scheduledDay;
    
    if (config.schedule.type === 'monthly') {
      scheduledDay = 1; // First of the month by default
    } else {
      scheduledDay = config.schedule.day || 1;
    }
    
    // Send if it's the scheduled day
    if (dayOfMonth === scheduledDay) {
      sendScheduledEmailReport();
    }
  } catch (error) {
    console.error("Error checking monthly email schedule: " + error.message);
  }
}

/**
 * Get list of recipients for sending email
 * @param {Object} config - Email configuration
 * @return {Array} List of recipient email addresses
 */
function getRecipientsListForSending(config) {
  const recipients = [];
  
  // Add configured recipients
  if (config.recipients && Array.isArray(config.recipients)) {
    recipients.push(...config.recipients.map(r => r.email));
  }
  
  // Add current user if includeMyself is enabled
  if (config.includeMyself) {
    const userEmail = UserService.getCurrentUserEmail();
    if (userEmail && !recipients.includes(userEmail)) {
      recipients.push(userEmail);
    }
  }
  
  // Ensure unique emails
  return [...new Set(recipients)];
}

/**
 * Generate PDF report
 * @param {Object} config - Email configuration
 * @param {Date} startDate - Start date for the report
 * @param {Date} endDate - End date for the report
 * @return {Blob} PDF blob or null if generation failed
 */
function generatePdfReport(config, startDate, endDate) {
  try {
    // Generate HTML content for the PDF
    const htmlContent = buildRichHtmlEmail(config, 
                                          MetricsService.getTeamMetrics(startDate, endDate),
                                          TeamService.getTeamMembersWithMetrics(startDate, endDate),
                                          startDate, endDate);
    
    // Format dates for display
    const dateRangeStr = `${formatDateForDisplay(startDate)}_${formatDateForDisplay(endDate)}`;
    
    // Create PDF from HTML
    const pdfBlob = HtmlService.createHtmlOutput(htmlContent)
      .getAs('application/pdf')
      .setName(`988_Crisis_Line_Metrics_${dateRangeStr}.pdf`);
    
    return pdfBlob;
  } catch (error) {
    console.error("Error generating PDF report: " + error.message);
    return null;
  }
}

/**
 * Generate spreadsheet report
 * @param {Object} config - Email configuration
 * @param {Date} startDate - Start date for the report
 * @param {Date} endDate - End date for the report
 * @return {Blob} Spreadsheet blob or null if generation failed
 */
function generateSpreadsheetReport(config, startDate, endDate) {
  try {
    // Create a new spreadsheet
    const spreadsheet = SpreadsheetApp.create(`988 Crisis Line Metrics ${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}`);
    
    // Team metrics sheet
    const teamSheet = spreadsheet.getActiveSheet();
    teamSheet.setName("Team Metrics");
    
    // Get team metrics data
    const teamMetrics = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Write headers
    teamSheet.getRange(1, 1, 1, 2).setValues([["Metric", "Value"]]);
    teamSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    
    // Write metrics data
    let row = 2;
    config.metrics.forEach(metric => {
      let value = null;
      
      switch (metric) {
        case 'Answer Rate':
          value = teamMetrics.answerRate;
          teamSheet.getRange(row, 2).setNumberFormat("0.0%");
          break;
        case 'Talk Time':
          value = teamMetrics.averageTalkTime;
          teamSheet.getRange(row, 2).setNumberFormat("0.0 \"min\"");
          break;
        case 'After Call Work':
          value = teamMetrics.acwAverage;
          teamSheet.getRange(row, 2).setNumberFormat("0.0 \"min\"");
          break;
        case 'On Queue Time':
          value = teamMetrics.onQueuePercentage;
          teamSheet.getRange(row, 2).setNumberFormat("0.0%");
          break;
        case 'Calls Handled':
          value = teamMetrics.callsAccepted;
          teamSheet.getRange(row, 2).setNumberFormat("#,##0");
          break;
        case 'Calls Offered':
          value = teamMetrics.callsOffered;
          teamSheet.getRange(row, 2).setNumberFormat("#,##0");
          break;
        // Add other metrics as needed...
        default:
          value = "N/A";
      }
      
      teamSheet.getRange(row, 1, 1, 2).setValues([[metric, value]]);
      row++;
    });
    
    // Auto-resize columns
    teamSheet.autoResizeColumns(1, 2);
    
    // Add individual metrics if needed
    if (config.reportType === 'individual' || config.reportType === 'comprehensive') {
      const teamMembers = TeamService.getTeamMembersWithMetrics(startDate, endDate);
      
      if (teamMembers && teamMembers.length > 0) {
        // Create new sheet for team members
        const membersSheet = spreadsheet.insertSheet("Individual Metrics");
        
        // Headers
        membersSheet.getRange(1, 1, 1, 5).setValues([
          ["Team Member", "Calls Handled", "Answer Rate", "Talk Time (min)", "Score"]
        ]);
        membersSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
        
        // Set data
        const memberData = teamMembers.map(member => [
          member.name,
          member.callsAccepted,
          member.answerRate,
          member.averageTalkTime,
          member.score
        ]);
        
        membersSheet.getRange(2, 1, memberData.length, 5).setValues(memberData);
        
        // Format the data
        membersSheet.getRange(2, 3, memberData.length, 1).setNumberFormat("0.0%"); // Answer Rate
        membersSheet.getRange(2, 4, memberData.length, 1).setNumberFormat("0.0"); // Talk Time
        membersSheet.getRange(2, 5, memberData.length, 1).setNumberFormat("0"); // Score
        
        // Auto-resize columns
        membersSheet.autoResizeColumns(1, 5);
      }
    }
    
    // Get the file as a blob
    const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    let blob;
    
    if (config.export && config.export.format === 'csv') {
      // Export as CSV
      blob = spreadsheetFile.getAs('text/csv')
        .setName(`988_Crisis_Line_Metrics_${formatDateForDisplay(startDate)}_${formatDateForDisplay(endDate)}.csv`);
    } else {
      // Export as Excel
      blob = spreadsheetFile.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        .setName(`988_Crisis_Line_Metrics_${formatDateForDisplay(startDate)}_${formatDateForDisplay(endDate)}.xlsx`);
    }
    
    // Delete the file from Drive after getting the blob
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    
    return blob;
  } catch (error) {
    console.error("Error generating spreadsheet report: " + error.message);
    return null;
  }
}

/**
 * Load user's export configuration
 * @return {Object} The export configuration
 */
function loadExportConfig() {
  try {
    // Get user-specific configuration
    const userEmail = UserService.getCurrentUserEmail();
    
    // Try to get stored config or use default
    let config = StorageService.getProperty(`${userEmail}_metrics_export_config`);
    if (!config) {
      return DEFAULT_CONFIG.export;
    }
    
    return JSON.parse(config);
  } catch (error) {
    console.error("Error loading export configuration: " + error.message);
    return DEFAULT_CONFIG.export;
  }
}

/**
 * Save user's export configuration
 * @param {Object} config - The export configuration to save
 * @return {Object} Result with success flag and message
 */
function saveExportConfiguration(config) {
  try {
    // Validate config
    validateExportConfig(config);
    
    // Get user-specific configuration
    const userEmail = UserService.getCurrentUserEmail();
    
    // Save to user's properties
    StorageService.setProperty(
      `${userEmail}_metrics_export_config`, 
      JSON.stringify(config)
    );
    
    // Set up export schedule trigger if enabled
    if (config.scheduleExport) {
      setupExportScheduleTrigger(config);
    } else {
      removeExportScheduleTrigger();
    }
    
    return {
      success: true,
      message: "Export configuration saved successfully"
    };
  } catch (error) {
    console.error("Error saving export configuration: " + error.message);
    
    return {
      success: false,
      message: "Error: " + error.message
    };
  }
}

/**
 * Validate export configuration
 * @param {Object} config - The configuration to validate
 * @throws {Error} If configuration is invalid
 */
function validateExportConfig(config) {
  // Check required fields
  if (!config.exportName) throw new Error("Export name is required");
  if (!config.exportType) throw new Error("Export type is required");
  if (!config.dateRange) throw new Error("Date range is required");
  if (!config.groupBy) throw new Error("Group by setting is required");
  if (!config.format) throw new Error("Export format is required");
  
  // Validate at least one field is selected
  if (!config.fields || !Array.isArray(config.fields) || config.fields.length === 0) {
    throw new Error("At least one field must be selected");
  }
}

/**
 * Set up export schedule trigger
 * @param {Object} config - Export configuration
 */
function setupExportScheduleTrigger(config) {
  try {
    // Remove any existing triggers
    removeExportScheduleTrigger();
    
    // For now, we'll use a weekly schedule for exports
    // Could be enhanced with custom scheduling similar to email reports
    ScriptApp.newTrigger('generateScheduledExport')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(1) // Run at 1am to not interfere with other processes
      .create();
  } catch (error) {
    console.error("Error setting up export schedule trigger: " + error.message);
    throw error;
  }
}

/**
 * Remove existing export schedule triggers
 */
function removeExportScheduleTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'generateScheduledExport') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Generate export now
 * @param {Object} [config] - Export configuration (optional, uses stored config if not provided)
 * @return {Object} Result with success flag, message, and file ID
 */
function generateExportNow(config = null) {
  try {
    // Use provided config or load from storage
    const exportConfig = config || loadExportConfig();
    
    // Get date range
    const endDate = new Date();
    let startDate;
    
    if (exportConfig.dateRange === 'custom' && exportConfig.startDate && exportConfig.endDate) {
      startDate = new Date(exportConfig.startDate);
      endDate = new Date(exportConfig.endDate);
    } else {
      startDate = getStartDateFromConfig(exportConfig.dateRange, endDate);
    }
    
    // Generate export based on format
    let fileId, fileName;
    
    if (exportConfig.format === 'pdf') {
      // Generate PDF report
      const result = generateMetricsPdfExport(exportConfig, startDate, endDate);
      fileId = result.fileId;
      fileName = result.fileName;
    } else if (exportConfig.format === 'csv') {
      // Generate CSV export
      const result = generateMetricsCsvExport(exportConfig, startDate, endDate);
      fileId = result.fileId;
      fileName = result.fileName;
    } else {
      // Generate spreadsheet export (default)
      const result = generateMetricsSpreadsheetExport(exportConfig, startDate, endDate);
      fileId = result.fileId;
      fileName = result.fileName;
    }
    
    return {
      success: true,
      message: `Export "${fileName}" generated successfully`,
      fileId: fileId,
      fileName: fileName,
      format: exportConfig.format
    };
  } catch (error) {
    console.error("Error generating export: " + error.message);
    
    return {
      success: false,
      message: "Error generating export: " + error.message
    };
  }
}

/**
 * Generate metrics PDF export
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Object} Result with fileId and fileName
 */
function generateMetricsPdfExport(config, startDate, endDate) {
  try {
    // Create PDF configuration based on export config
    const pdfConfig = {
      reportName: config.exportName,
      reportType: config.exportType,
      reportFormat: 'rich',
      metrics: config.fields,
      charts: config.includeCharts ? ['Call Volume', 'Answer Rate'] : [],
      introMessage: `This report contains metrics data for the period from ${formatDateForDisplay(startDate)} to ${formatDateForDisplay(endDate)}.`
    };
    
    // Generate PDF content
    const htmlContent = buildRichHtmlEmail(pdfConfig, 
                                          MetricsService.getTeamMetrics(startDate, endDate),
                                          TeamService.getTeamMembersWithMetrics(startDate, endDate),
                                          startDate, endDate);
    
    // Format dates for filename
    const dateStr = `${startDate.toISOString().split('T')[0]}_to_${endDate.toISOString().split('T')[0]}`;
    const fileName = `${config.exportName.replace(/\s+/g, '_')}_${dateStr}.pdf`;
    
    // Create PDF
    const pdfBlob = HtmlService.createHtmlOutput(htmlContent)
      .getAs('application/pdf')
      .setName(fileName);
    
    // Save to Drive
    const userFolder = getUserExportFolder();
    const pdfFile = userFolder.createFile(pdfBlob);
    
    return {
      fileId: pdfFile.getId(),
      fileName: fileName
    };
  } catch (error) {
    console.error("Error generating PDF export: " + error.message);
    throw error;
  }
}

/**
 * Generate metrics CSV export
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Object} Result with fileId and fileName
 */
function generateMetricsCsvExport(config, startDate, endDate) {
  try {
    // Format dates for filename
    const dateStr = `${startDate.toISOString().split('T')[0]}_to_${endDate.toISOString().split('T')[0]}`;
    const fileName = `${config.exportName.replace(/\s+/g, '_')}_${dateStr}.csv`;
    
    // Create CSV content based on export type
    let csvContent = "";
    
    if (config.exportType === 'team') {
      csvContent = generateTeamMetricsCsv(config, startDate, endDate);
    } else if (config.exportType === 'individual') {
      csvContent = generateIndividualMetricsCsv(config, startDate, endDate);
    } else if (config.exportType === 'raw') {
      csvContent = generateRawDataCsv(config, startDate, endDate);
    } else {
      // Comprehensive - includes both team and individual metrics
      csvContent = generateComprehensiveMetricsCsv(config, startDate, endDate);
    }
    
    // Create blob and save to Drive
    const blob = Utilities.newBlob(csvContent, "text/csv", fileName);
    const userFolder = getUserExportFolder();
    const file = userFolder.createFile(blob);
    
    return {
      fileId: file.getId(),
      fileName: fileName
    };
  } catch (error) {
    console.error("Error generating CSV export: " + error.message);
    logExportError(error, config);
    throw error;
  }
}

/**
 * Generate CSV content for team metrics
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} CSV content
 */
function generateTeamMetricsCsv(config, startDate, endDate) {
  try {
    // Get metrics data
    const metricsData = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Build CSV content
    let csv = `"988 Crisis Line Team Metrics Report"\n`;
    csv += `"Date Range","${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}"\n`;
    csv += `"Generated On","${formatDateForDisplay(new Date())} ${new Date().toLocaleTimeString()}"\n\n`;
    
    // Summary metrics
    csv += `"Summary Metrics"\n`;
    
    // Headers - Metric, Value
    csv += `"Metric","Value"\n`;
    
    // Add selected fields
    for (const field of config.fields) {
      const value = getMetricValueByName(field, metricsData);
      const formattedValue = formatMetricValueForCsv(field, value);
      csv += `"${field}","${formattedValue}"\n`;
    }
    
    // If we have daily metrics and the grouping is by day, add them
    if (metricsData.dailyMetrics && metricsData.dailyMetrics.length > 0 && config.groupBy === 'day') {
      csv += `\n"Daily Breakdown"\n`;
      
      // Headers - Date and all selected fields
      let headers = [`"Date"`];
      config.fields.forEach(field => {
        headers.push(`"${field}"`);
      });
      csv += headers.join(',') + '\n';
      
      // Add data rows
      metricsData.dailyMetrics.forEach(day => {
        let row = [`"${formatDateForDisplay(new Date(day.date))}"`];
        
        config.fields.forEach(field => {
          const value = getMetricValueByName(field, day);
          const formattedValue = formatMetricValueForCsv(field, value);
          row.push(`"${formattedValue}"`);
        });
        
        csv += row.join(',') + '\n';
      });
    }
    
    // If grouping by week and we have daily data, aggregate and add weekly breakdown
    if (metricsData.dailyMetrics && metricsData.dailyMetrics.length > 0 && config.groupBy === 'week') {
      csv += `\n"Weekly Breakdown"\n`;
      
      // Group metrics by week
      const weeklyMetrics = aggregateMetricsByWeek(metricsData.dailyMetrics);
      
      // Headers - Week and all selected fields
      let headers = [`"Week Starting"`];
      config.fields.forEach(field => {
        headers.push(`"${field}"`);
      });
      csv += headers.join(',') + '\n';
      
      // Add data rows for each week
      Object.keys(weeklyMetrics).sort().forEach(weekStart => {
        const weekData = weeklyMetrics[weekStart];
        let row = [`"${formatDateForDisplay(new Date(weekStart))}"`];
        
        config.fields.forEach(field => {
          const value = getMetricValueByName(field, weekData);
          const formattedValue = formatMetricValueForCsv(field, value);
          row.push(`"${formattedValue}"`);
        });
        
        csv += row.join(',') + '\n';
      });
    }
    
    // If grouping by month and we have daily data, aggregate and add monthly breakdown
    if (metricsData.dailyMetrics && metricsData.dailyMetrics.length > 0 && config.groupBy === 'month') {
      csv += `\n"Monthly Breakdown"\n`;
      
      // Group metrics by month
      const monthlyMetrics = aggregateMetricsByMonth(metricsData.dailyMetrics);
      
      // Headers - Month and all selected fields
      let headers = [`"Month"`];
      config.fields.forEach(field => {
        headers.push(`"${field}"`);
      });
      csv += headers.join(',') + '\n';
      
      // Add data rows for each month
      Object.keys(monthlyMetrics).sort().forEach(monthStart => {
        const monthData = monthlyMetrics[monthStart];
        let row = [`"${formatMonthForDisplay(new Date(monthStart))}"`];
        
        config.fields.forEach(field => {
          const value = getMetricValueByName(field, monthData);
          const formattedValue = formatMetricValueForCsv(field, value);
          row.push(`"${formattedValue}"`);
        });
        
        csv += row.join(',') + '\n';
      });
    }
    
    return csv;
  } catch (error) {
    console.error("Error generating team metrics CSV content: " + error.message);
    throw error;
  }
}

/**
 * Generate CSV content for individual metrics
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} CSV content
 */
function generateIndividualMetricsCsv(config, startDate, endDate) {
  try {
    // Get team members data
    const teamMembers = TeamService.getTeamMembersWithMetrics(startDate, endDate);
    
    if (!teamMembers || teamMembers.length === 0) {
      return `"No team member data available for the selected date range."`;
    }
    
    // Build CSV content
    let csv = `"988 Crisis Line Individual Metrics Report"\n`;
    csv += `"Date Range","${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}"\n`;
    csv += `"Generated On","${formatDateForDisplay(new Date())} ${new Date().toLocaleTimeString()}"\n\n`;
    
    // Summary - all team members
    csv += `"Team Member Performance Summary"\n`;
    
    // Headers
    let headers = [`"Team Member"`];
    config.fields.forEach(field => {
      headers.push(`"${field}"`);
    });
    headers.push(`"Overall Score"`);
    csv += headers.join(',') + '\n';
    
    // Add data for each team member
    teamMembers.forEach(member => {
      let row = [`"${member.name}"`];
      
      // Add values for configured fields
      config.fields.forEach(field => {
        const value = getMetricValueByName(field, member);
        const formattedValue = formatMetricValueForCsv(field, value);
        row.push(`"${formattedValue}"`);
      });
      
      // Add overall score
      row.push(`"${Math.round(member.score)}"`);
      
      csv += row.join(',') + '\n';
    });
    
    // If including comparisons, add team averages
    if (config.includeComparisons) {
      // Calculate team averages
      const teamAverages = calculateTeamAverages(teamMembers);
      
      // Add team average row
      let avgRow = [`"TEAM AVERAGE"`];
      
      // Add values for configured fields
      config.fields.forEach(field => {
        const value = getMetricValueByName(field, teamAverages);
        const formattedValue = formatMetricValueForCsv(field, value);
        avgRow.push(`"${formattedValue}"`);
      });
      
      // Add average score
      avgRow.push(`"${Math.round(teamAverages.score)}"`);
      
      csv += avgRow.join(',') + '\n';
    }
    
    // If grouping by team_member, include detailed history for each
    if (config.groupBy === 'team_member') {
      // For each team member, add their metrics history
      teamMembers.forEach(member => {
        if (member.metricsHistory && member.metricsHistory.length > 0) {
          csv += `\n"${member.name} - Daily Metrics"\n`;
          
          // Headers for daily metrics
          let historyHeaders = [`"Date"`];
          config.fields.forEach(field => {
            historyHeaders.push(`"${field}"`);
          });
          csv += historyHeaders.join(',') + '\n';
          
          // Add daily data rows
          member.metricsHistory.forEach(day => {
            let row = [`"${formatDateForDisplay(new Date(day.date))}"`];
            
            config.fields.forEach(field => {
              const value = getMetricValueByName(field, day);
              const formattedValue = formatMetricValueForCsv(field, value);
              row.push(`"${formattedValue}"`);
            });
            
            csv += row.join(',') + '\n';
          });
        }
      });
    }
    
    return csv;
  } catch (error) {
    console.error("Error generating individual metrics CSV content: " + error.message);
    throw error;
  }
}

/**
 * Generate CSV content for raw data export
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} CSV content
 */
function generateRawDataCsv(config, startDate, endDate) {
  try {
    // Get raw call data
    const rawData = MetricsService.getRawCallData(startDate, endDate);
    
    if (!rawData || rawData.length === 0) {
      return `"No raw call data available for the selected date range."`;
    }
    
    // Apply anonymization if needed
    const processedData = config.anonymize ? anonymizeRawData(rawData) : rawData;
    
    // Build CSV content
    let csv = `"988 Crisis Line Raw Call Data Export"\n`;
    csv += `"Date Range","${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}"\n`;
    csv += `"Generated On","${formatDateForDisplay(new Date())} ${new Date().toLocaleTimeString()}"\n`;
    csv += `"Records","${processedData.length}"\n\n`;
    
    // Determine headers from the first data entry
    const headerKeys = Object.keys(processedData[0] || {});
    
    // Add headers
    const headers = headerKeys.map(key => `"${formatHeaderName(key)}"`);
    csv += headers.join(',') + '\n';
    
    // Add data rows
    processedData.forEach(call => {
      const row = headerKeys.map(key => {
        // Format based on field type
        if (key.includes('date') || key.includes('time')) {
          // Format dates and times
          return `"${formatDateTimeForCsv(call[key])}"`;
        } else if (typeof call[key] === 'number') {
          return call[key];
        } else {
          // Ensure strings are properly escaped for CSV
          return `"${(call[key] || '').toString().replace(/"/g, '""')}"`;
        }
      });
      
      csv += row.join(',') + '\n';
    });
    
    return csv;
  } catch (error) {
    console.error("Error generating raw data CSV content: " + error.message);
    throw error;
  }
}

/**
 * Generate comprehensive CSV export with all metrics
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {string} CSV content
 */
function generateComprehensiveMetricsCsv(config, startDate, endDate) {
  try {
    // Get both team and individual metrics
    let csv = `"988 Crisis Line Comprehensive Metrics Report"\n`;
    csv += `"Date Range","${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}"\n`;
    csv += `"Generated On","${formatDateForDisplay(new Date())} ${new Date().toLocaleTimeString()}"\n\n`;
    
    // Get team metrics
    const teamConfig = { ...config, exportType: 'team' };
    const teamCsv = generateTeamMetricsCsv(teamConfig, startDate, endDate);
    
    // Strip the first few header lines from team CSV since we already have headers
    const teamLines = teamCsv.split('\n');
    const teamContent = teamLines.slice(4).join('\n'); // Skip the first 4 lines
    
    // Get individual metrics
    const individualConfig = { ...config, exportType: 'individual' };
    const individualCsv = generateIndividualMetricsCsv(individualConfig, startDate, endDate);
    
    // Strip the first few header lines from individual CSV
    const individualLines = individualCsv.split('\n');
    const individualContent = individualLines.slice(4).join('\n'); // Skip the first 4 lines
    
    // Combine everything
    csv += teamContent + '\n\n' + individualContent;
    
    return csv;
  } catch (error) {
    console.error("Error generating comprehensive metrics CSV content: " + error.message);
    throw error;
  }
}

/**
 * Generate spreadsheet export for metrics
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Object} Result with fileId and fileName
 */
function generateMetricsSpreadsheetExport(config, startDate, endDate) {
  try {
    // Format dates for filename
    const dateStr = `${startDate.toISOString().split('T')[0]}_to_${endDate.toISOString().split('T')[0]}`;
    const fileName = `${config.exportName.replace(/\s+/g, '_')}_${dateStr}.xlsx`;
    
    // Create a new spreadsheet
    const spreadsheet = SpreadsheetApp.create(fileName.replace('.xlsx', ''));
    
    // Generate content based on export type
    if (config.exportType === 'team') {
      exportTeamMetrics(spreadsheet, config, startDate, endDate);
    } else if (config.exportType === 'individual') {
      exportIndividualMetrics(spreadsheet, config, startDate, endDate);
    } else if (config.exportType === 'raw') {
      exportRawData(spreadsheet, config, startDate, endDate);
    } else {
      // Comprehensive - export all types
      exportComprehensiveMetrics(spreadsheet, config, startDate, endDate);
    }
    
    // Add charts if requested
    if (config.includeCharts && config.exportType !== 'raw') {
      addChartsToSpreadsheet(spreadsheet, config, startDate, endDate);
    }
    
    // Convert to Excel format
    const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    const excelBlob = spreadsheetFile.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').setName(fileName);
    
    // Save to user's folder
    const userFolder = getUserExportFolder();
    const excelFile = userFolder.createFile(excelBlob);
    
    // Delete the original Google Sheets file
    spreadsheetFile.setTrashed(true);
    
    return {
      fileId: excelFile.getId(),
      fileName: fileName
    };
  } catch (error) {
    console.error("Error generating spreadsheet export: " + error.message);
    logExportError(error, config);
    throw error;
  }
}

/**
 * Export team metrics to a spreadsheet
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function exportTeamMetrics(spreadsheet, config, startDate, endDate) {
  try {
    // Get metrics data
    const metricsData = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Use the active sheet for summary
    const summarySheet = spreadsheet.getActiveSheet();
    summarySheet.setName("Summary");
    
    // Add title and date range
    summarySheet.getRange(1, 1).setValue("988 Crisis Line Team Metrics Report");
    summarySheet.getRange(1, 1, 1, 5).merge();
    summarySheet.getRange(1, 1).setFontWeight("bold").setFontSize(14);
    
    summarySheet.getRange(2, 1).setValue(`Date Range: ${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}`);
    summarySheet.getRange(2, 1, 1, 5).merge();
    
    // Add metrics
    summarySheet.getRange(4, 1).setValue("Metric");
    summarySheet.getRange(4, 2).setValue("Value");
    summarySheet.getRange(4, 1, 1, 2).setFontWeight("bold").setBackground("#f1f3f4");
    
    // Add selected fields
    for (let i = 0; i < config.fields.length; i++) {
      const field = config.fields[i];
      const row = i + 5;
      
      const value = getMetricValueByName(field, metricsData);
      
      summarySheet.getRange(row, 1).setValue(field);
      summarySheet.getRange(row, 2).setValue(value);
      
      // Apply formatting
      applyMetricFormatting(summarySheet, row, 2, field, value);
    }
    
    // Add daily metrics if available
    if (metricsData.dailyMetrics && metricsData.dailyMetrics.length > 0) {
      // Create new sheet for daily metrics
      const dailySheet = spreadsheet.insertSheet("Daily Metrics");
      
      // Headers
      const headers = ["Date"].concat(config.fields);
      dailySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      dailySheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f3f4");
      
      // Prepare data rows
      const data = metricsData.dailyMetrics.map(day => {
        const row = [new Date(day.date)];
        
        // Add values for each selected field
        config.fields.forEach(field => {
          const value = getMetricValueByName(field, day);
          row.push(value);
        });
        
        return row;
      });
      
      // Write data
      if (data.length > 0) {
        dailySheet.getRange(2, 1, data.length, headers.length).setValues(data);
        
        // Format date column
        dailySheet.getRange(2, 1, data.length, 1).setNumberFormat("yyyy-mm-dd");
        
        // Apply formatting to each field
        for (let i = 0; i < config.fields.length; i++) {
          const field = config.fields[i];
          const col = i + 2;
          
          applyColumnFormatting(dailySheet, 2, col, data.length, 1, field);
        }
        
        // Auto-resize columns
        dailySheet.autoResizeColumns(1, headers.length);
        
        // Add weekly aggregation if requested
        if (config.groupBy === 'week') {
          exportWeeklyAggregation(spreadsheet, config, metricsData.dailyMetrics);
        }
        
        // Add monthly aggregation if requested
        if (config.groupBy === 'month') {
          exportMonthlyAggregation(spreadsheet, config, metricsData.dailyMetrics);
        }
      }
    }
    
    // Auto-resize summary sheet columns
    summarySheet.autoResizeColumns(1, 2);
  } catch (error) {
    console.error("Error exporting team metrics to spreadsheet: " + error.message);
    throw error;
  }
}

/**
 * Export individual metrics to spreadsheet
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function exportIndividualMetrics(spreadsheet, config, startDate, endDate) {
  try {
    // Get team members data
    const teamMembers = TeamService.getTeamMembersWithMetrics(startDate, endDate);
    
    if (!teamMembers || teamMembers.length === 0) {
      // Use the active sheet for message
      const sheet = spreadsheet.getActiveSheet();
      sheet.setName("Team Performance");
      sheet.getRange(1, 1).setValue("No team member data available for the selected date range");
      return;
    }
    
    // Use the active sheet for team member summary
    const summarySheet = spreadsheet.getActiveSheet();
    summarySheet.setName("Team Performance");
    
    // Add title and date range
    summarySheet.getRange(1, 1).setValue("988 Crisis Line Team Member Performance");
    summarySheet.getRange(1, 1, 1, 8).merge();
    summarySheet.getRange(1, 1).setFontWeight("bold").setFontSize(14);
    
    summarySheet.getRange(2, 1).setValue(`Date Range: ${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}`);
    summarySheet.getRange(2, 1, 1, 8).merge();
    
    // Create headers based on selected fields
    const headers = ["Team Member"];
    config.fields.forEach(field => headers.push(field));
    headers.push("Score");
    
    summarySheet.getRange(4, 1, 1, headers.length).setValues([headers]);
    summarySheet.getRange(4, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f3f4");
    
    // Sort team members by score
    teamMembers.sort((a, b) => b.score - a.score);
    
    // Prepare team member rows
    const memberRows = teamMembers.map(member => {
      const row = [member.name];
      
      // Add values for each selected field
      config.fields.forEach(field => {
        const value = getMetricValueByName(field, member);
        row.push(value);
      });
      
      // Add score
      row.push(member.score);
      
      return row;
    });
    
    // Write data
    if (memberRows.length > 0) {
      summarySheet.getRange(5, 1, memberRows.length, headers.length).setValues(memberRows);
      
      // Apply formatting to each field
      for (let i = 0; i < config.fields.length; i++) {
        const field = config.fields[i];
        const col = i + 2;
        
        applyColumnFormatting(summarySheet, 5, col, memberRows.length, 1, field);
      }
      
      // Format score column
      summarySheet.getRange(5, headers.length, memberRows.length, 1).setNumberFormat("0.0");
    }
    
    // Create sheet for each team member if grouping by team_member
    if (config.groupBy === 'team_member') {
      // Maximum of 10 team members to avoid creating too many sheets
      const topTeamMembers = teamMembers
        .filter(member => member.metricsHistory && member.metricsHistory.length > 0)
        .slice(0, 10);
      
      // Create sheet for each team member
      topTeamMembers.forEach(member => {
        // Create sheet with safe name (max 31 chars, no special chars)
        const safeName = member.name.replace(/[^\w\s]/gi, '').substring(0, 30);
        const memberSheet = spreadsheet.insertSheet(safeName);
        
        // Set headers
        const headers = ['Date'].concat(config.fields);
        memberSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        memberSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
        
        if (!member.metricsHistory || member.metricsHistory.length === 0) {
          memberSheet.getRange(2, 1).setValue("No metrics history available for this team member");
          return; // Skip to next team member
        }
        
        // Prepare data rows
        const data = member.metricsHistory.map(day => {
          const row = [new Date(day.date)]; // Format as date for spreadsheet
          
          config.fields.forEach(field => {
            const value = getMetricValueByName(field, day);
            row.push(value);
          });
          
          return row;
        });
        
        // Write data to sheet
        memberSheet.getRange(2, 1, data.length, headers.length).setValues(data);
        
        // Format the data
        // Date column
        memberSheet.getRange(2, 1, data.length, 1).setNumberFormat("yyyy-mm-dd");
        
        // Format other columns based on field type
        for (let i = 0; i < config.fields.length; i++) {
          const field = config.fields[i];
          const col = i + 2;
          
          applyColumnFormatting(memberSheet, 2, col, data.length, 1, field);
        }
        
        // Auto-resize columns
        memberSheet.autoResizeColumns(1, headers.length);
      });
    }
    
    // Auto-resize columns in summary sheet
    summarySheet.autoResizeColumns(1, headers.length);
  } catch (error) {
    console.error("Error exporting individual metrics to spreadsheet: " + error.message);
    throw error;
  }
}

/**
 * Export team member summary to a sheet
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {Array} teamMembers - Team members with metrics
 * @param {Object} config - Export configuration
 */
function exportTeamMemberSummary(spreadsheet, teamMembers, config) {
  try {
    // Create a new sheet
    const sheet = spreadsheet.insertSheet("Team Member Summary");
    
    // Set headers
    const headers = ["Team Member", "Calls Accepted", "Answer Rate", "Talk Time", "ACW", "On Queue %", "Goal Compliance", "Score"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    
    if (!teamMembers || teamMembers.length === 0) {
      sheet.getRange(2, 1).setValue("No team member data available for the selected period");
      sheet.autoResizeColumns(1, headers.length);
      return;
    }
    
    // Sort team members by score (descending)
    teamMembers.sort((a, b) => b.score - a.score);
    
    // Prepare team member data
    const teamData = teamMembers.map(member => {
      // Calculate goal compliance as percentage of goals met
      const answerRateGoal = member.answerRate >= 0.95;
      const talkTimeGoal = member.averageTalkTime >= 15 && member.averageTalkTime <= 20;
      const acwGoal = member.acwAverage <= 5;
      const onQueueGoal = member.onQueuePercentage >= 0.65;
      
      const goalsCompliant = [answerRateGoal, talkTimeGoal, acwGoal, onQueueGoal];
      const goalsMet = goalsCompliant.filter(goal => goal).length;
      const goalCompliance = goalsMet / goalsCompliant.length;
      
      return [
        member.name,
        member.callsAccepted,
        member.answerRate,
        member.averageTalkTime,
        member.acwAverage,
        member.onQueuePercentage,
        goalCompliance,
        member.score
      ];
    });
    
    // Write data to sheet
    sheet.getRange(2, 1, teamData.length, headers.length).setValues(teamData);
    
    // Format the data
    sheet.getRange(2, 3, teamData.length, 1).setNumberFormat("0.0%"); // Answer Rate
    sheet.getRange(2, 4, teamData.length, 1).setNumberFormat("0.0"); // Talk Time
    sheet.getRange(2, 5, teamData.length, 1).setNumberFormat("0.0"); // ACW
    sheet.getRange(2, 6, teamData.length, 1).setNumberFormat("0.0%"); // On Queue %
    sheet.getRange(2, 7, teamData.length, 1).setNumberFormat("0%"); // Goal Compliance
    sheet.getRange(2, 8, teamData.length, 1).setNumberFormat("0.0"); // Score
    
    // Add conditional formatting
    const answerRateRange = sheet.getRange(2, 3, teamData.length, 1);
    const talkTimeRange = sheet.getRange(2, 4, teamData.length, 1);
    const acwRange = sheet.getRange(2, 5, teamData.length, 1);
    const onQueueRange = sheet.getRange(2, 6, teamData.length, 1);
    const scoreRange = sheet.getRange(2, 8, teamData.length, 1);
    
    // Answer Rate: ≥95% is good (green), ≥90% is average (yellow), <90% is poor (red)
    answerRateRange.setDataValidation(null); // Clear any existing rules
    answerRateRange.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.95)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([answerRateRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.9)
        .whenNumberLessThan(0.95)
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([answerRateRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0.9)
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([answerRateRange])
        .build()
    ]);
    
    // Talk Time: 15-20 min is good (green), 13-15 or 20-22 is average (yellow), <13 or >22 is poor (red)
    talkTimeRange.setDataValidation(null); // Clear any existing rules
    talkTimeRange.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(15, 20)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([talkTimeRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=OR(AND(D2>=13,D2<15),AND(D2>20,D2<=22))')
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([talkTimeRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=OR(D2<13,D2>22)')
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([talkTimeRange])
        .build()
    ]);
    
    // ACW Time: ≤5 min is good (green), 5-7 min is average (yellow), >7 min is poor (red)
    acwRange.setDataValidation(null); // Clear any existing rules
    acwRange.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThanOrEqualTo(5)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([acwRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(5)
        .whenNumberLessThanOrEqualTo(7)
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([acwRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(7)
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([acwRange])
        .build()
    ]);
    
    // On Queue %: ≥65% is good (green), 60-65% is average (yellow), <60% is poor (red)
    onQueueRange.setDataValidation(null); // Clear any existing rules
    onQueueRange.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.65)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([onQueueRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.6)
        .whenNumberLessThan(0.65)
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([onQueueRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0.6)
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([onQueueRange])
        .build()
    ]);
    
    // Score: ≥90 is excellent (dark green), ≥80 is good (green), ≥70 is average (yellow), ≥60 is fair (orange), <60 is poor (red)
    scoreRange.setDataValidation(null); // Clear any existing rules
    scoreRange.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(90)
        .setBackground('#0F9D58')
        .setFontColor('#FFFFFF')
        .setBold(true)
        .setRanges([scoreRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(80)
        .whenNumberLessThan(90)
        .setBackground('#34A853')
        .setFontColor('#FFFFFF')
        .setBold(true)
        .setRanges([scoreRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(70)
        .whenNumberLessThan(80)
        .setBackground('#FBBC05')
        .setFontColor('#202124')
        .setBold(true)
        .setRanges([scoreRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(60)
        .whenNumberLessThan(70)
        .setBackground('#F9AB00')
        .setFontColor('#FFFFFF')
        .setBold(true)
        .setRanges([scoreRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(60)
        .setBackground('#EA4335')
        .setFontColor('#FFFFFF')
        .setBold(true)
        .setRanges([scoreRange])
        .build()
    ]);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
  } catch (error) {
    console.error("Error exporting team member summary: " + error.message);
    throw error;
  }
}

/**
 * Export all team member sheets to a spreadsheet
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function exportAllTeamMemberSheets(spreadsheet, config, startDate, endDate) {
  try {
    // Get team members with metrics history
    const teamMembers = TeamService.getTeamMembersWithMetricsHistory(startDate, endDate);
    
    if (!teamMembers || teamMembers.length === 0) {
      return; // No team members to export
    }
    
    // Maximum of 10 team members to avoid creating too many sheets
    const topTeamMembers = teamMembers
      .filter(member => member.metricsHistory && member.metricsHistory.length > 0)
      .slice(0, 10);
    
    // Create sheet for each team member
    topTeamMembers.forEach(member => {
      // Create sheet with safe name (max 31 chars, no special chars)
      const safeName = member.name.replace(/[^\w\s]/gi, '').substring(0, 30);
      const memberSheet = spreadsheet.insertSheet(safeName);
      
      // Set headers
      const headers = ['Date'].concat(config.fields);
      memberSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      memberSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      
      // Check if this team member has metrics history
      if (!member.metricsHistory || member.metricsHistory.length === 0) {
        memberSheet.getRange(2, 1).setValue("No metrics history available for this team member");
        memberSheet.getRange(2, 1, 1, headers.length).merge();
        memberSheet.getRange(2, 1).setHorizontalAlignment("center");
        memberSheet.getRange(2, 1).setFontStyle("italic");
        return; // Skip to next member
      }
      
      // Prepare data rows
      const data = member.metricsHistory.map(day => {
        const row = [new Date(day.date)]; // Format as date for spreadsheet
        
        config.fields.forEach(field => {
          let value = null;
          
          switch (field) {
            case 'Answer Rate':
              value = day.answerRate;
              break;
            case 'Talk Time':
              value = day.averageTalkTime;
              break;
            case 'After Call Work':
              value = day.acwAverage;
              break;
            case 'On Queue Time':
              value = day.onQueuePercentage;
              break;
            case 'Off Queue Time':
              value = day.offQueuePercentage;
              break;
            case 'Calls Accepted':
              value = day.callsAccepted;
              break;
            case 'Interacting Time':
              value = day.interactingTime;
              break;
            // Add other metrics as needed
            default:
              value = "N/A";
          }
          
          row.push(value);
        });
        
        return row;
      });
      
      // Write data to sheet
      memberSheet.getRange(2, 1, data.length, headers.length).setValues(data);
      
      // Format the data
      // Date column
      memberSheet.getRange(2, 1, data.length, 1).setNumberFormat("yyyy-mm-dd");
      
      // Other columns based on field type
      const formats = {
        'Answer Rate': "0.00%",
        'On Queue Time': "0.00%",
        'Off Queue Time': "0.00%",
        'Talk Time': "0.0 \"min\"",
        'After Call Work': "0.0 \"min\"",
        'Interacting Time': "0.0 \"hrs\"",
        'Calls Accepted': "#,##0"
      };
      
      config.fields.forEach((field, index) => {
        if (formats[field]) {
          memberSheet.getRange(2, index + 2, data.length, 1).setNumberFormat(formats[field]);
        }
      });
      
      // Auto-resize columns
      memberSheet.autoResizeColumns(1, headers.length);
    });
  } catch (error) {
    console.error("Error exporting all team member sheets: " + error.message);
    logExportError(error, { type: "team_member_sheets", startDate, endDate });
  }
}
/**
 * Get metric value by field name from data object
 * @param {string} fieldName - Field name (display name)
 * @param {Object} data - Data object containing metrics
 * @return {any} Metric value or null if not found
 */
function getMetricValueByName(fieldName, data) {
  if (!data) return null;
  
  switch (fieldName) {
    case 'Answer Rate':
      return data.answerRate || data.answer_rate || 0;
    case 'Talk Time':
      return data.averageTalkTime || data.average_talk_time || data.talk_time || 0;
    case 'After Call Work':
      return data.acwAverage || data.acw_average || data.acw || 0;
    case 'On Queue Time':
      return data.onQueuePercentage || data.on_queue_percentage || data.on_queue || 0;
    case 'Off Queue Time':
      return data.offQueuePercentage || data.off_queue_percentage || data.off_queue || 0;
    case 'Calls Offered':
      return data.callsOffered || data.calls_offered || 0;
    case 'Calls Accepted':
      return data.callsAccepted || data.calls_accepted || data.calls_handled || 0;
    case 'Interacting Time':
      return data.interactingTimePercentage || data.interacting_time_percentage || data.interactingTime || 0;
    case 'Staff Count':
      return data.staffCount || data.staff_count || 0;
    default:
      // Try various property name formats
      const camelCase = toCamelCase(fieldName);
      const snakeCase = toSnakeCase(fieldName);
      const lowerCase = fieldName.toLowerCase().replace(/\s+/g, '');
      
      return data[camelCase] || data[snakeCase] || data[lowerCase] || null;
  }
}

/**
 * Format metric value for CSV export
 * @param {string} fieldName - Field name
 * @param {any} value - Metric value
 * @return {string} Formatted value
 */
function formatMetricValueForCsv(fieldName, value) {
  if (value === null || value === undefined) {
    return 'N/A';
  }
  
  switch (fieldName) {
    case 'Answer Rate':
    case 'On Queue Time':
    case 'Off Queue Time':
    case 'Interacting Time':
      // Format percentages
      return isNaN(value) ? 'N/A' : (value * 100).toFixed(1) + '%';
    case 'Talk Time':
    case 'After Call Work':
      // Format times
      return isNaN(value) ? 'N/A' : value.toFixed(1) + ' min';
    case 'Calls Offered':
    case 'Calls Accepted':
    case 'Staff Count':
      // Format integers
      return isNaN(value) ? 'N/A' : Math.round(value).toString();
    default:
      // Handle unknown field types
      if (typeof value === 'number') {
        return value % 1 === 0 ? value.toString() : value.toFixed(2);
      } else {
        return value.toString();
      }
  }
}

/**
 * Format header name for export
 * @param {string} key - Object property key
 * @return {string} Formatted header name
 */
function formatHeaderName(key) {
  // Convert camelCase or snake_case to Title Case
  return key
    .replace(/_/g, ' ')
    .replace(/([A-Z])/g, ' $1')
    .replace(/^./, str => str.toUpperCase())
    .trim();
}

/**
 * Format date/time for CSV export
 * @param {string|Date} dateTime - Date or datetime string/object
 * @return {string} Formatted date/time string
 */
function formatDateTimeForCsv(dateTime) {
  if (!dateTime) return '';
  
  try {
    const date = new Date(dateTime);
    
    if (isNaN(date.getTime())) {
      return dateTime.toString();
    }
    
    // Check if time component is non-zero
    if (date.getHours() === 0 && date.getMinutes() === 0 && date.getSeconds() === 0) {
      // Format as date only
      return date.toLocaleDateString();
    } else {
      // Format as date and time
      return `${date.toLocaleDateString()} ${date.toLocaleTimeString()}`;
    }
  } catch (e) {
    return dateTime.toString();
  }
}

/**
 * Anonymize raw call data
 * @param {Array} rawData - Raw call data
 * @return {Array} Anonymized data
 */
function anonymizeRawData(rawData) {
  if (!rawData || !Array.isArray(rawData)) return [];
  
  // Fields to anonymize
  const sensitiveFields = [
    'caller_id', 'caller_name', 'phone_number', 'contact_info', 
    'email', 'address', 'notes', 'zip_code', 'location'
  ];
  
  return rawData.map(record => {
    // Create a copy of the record
    const anonymized = {...record};
    
    // Anonymize sensitive fields
    Object.keys(anonymized).forEach(key => {
      const lowerKey = key.toLowerCase();
      
      // Check if this is a sensitive field
      if (sensitiveFields.some(field => lowerKey.includes(field)) ||
          (lowerKey.includes('id') && !lowerKey.includes('call_id'))) {
        
        if (typeof anonymized[key] === 'string') {
          // Replace with anonymized version
          if (lowerKey.includes('phone') || lowerKey.includes('caller_id')) {
            anonymized[key] = 'XXX-XXX-XXXX';
          } else if (lowerKey.includes('email')) {
            anonymized[key] = 'user@redacted.com';
          } else if (lowerKey.includes('name')) {
            anonymized[key] = 'Redacted Name';
          } else if (lowerKey.includes('address')) {
            anonymized[key] = 'Redacted Address';
          } else if (lowerKey.includes('zip')) {
            anonymized[key] = 'XXXXX';
          } else if (lowerKey.includes('notes')) {
            anonymized[key] = 'Redacted notes';
          } else {
            anonymized[key] = '[REDACTED]';
          }
        }
      }
    });
    
    return anonymized;
  });
}

/**
 * Get user's export folder
 * @return {Folder} User's export folder in Drive
 */
function getUserExportFolder() {
  try {
    const userEmail = UserService.getCurrentUserEmail();
    const folderName = `988 Crisis Line Exports - ${userEmail}`;
    
    // Try to find existing folder
    const folderIterator = DriveApp.getFoldersByName(folderName);
    
    if (folderIterator.hasNext()) {
      return folderIterator.next();
    }
    
    // Create new folder if it doesn't exist
    return DriveApp.createFolder(folderName);
  } catch (error) {
    console.error("Error getting user export folder: " + error.message);
    // Fall back to root folder
    return DriveApp.getRootFolder();
  }
}

/**
 * Aggregate metrics by week
 * @param {Array} dailyMetrics - Daily metrics data
 * @return {Object} Weekly aggregated data
 */
function aggregateMetricsByWeek(dailyMetrics) {
  if (!dailyMetrics || !Array.isArray(dailyMetrics)) return {};
  
  const weeklyData = {};
  
  // Group by week
  dailyMetrics.forEach(day => {
    const date = new Date(day.date);
    
    // Get start of week (Sunday)
    const weekStart = new Date(date);
    weekStart.setDate(date.getDate() - date.getDay());
    weekStart.setHours(0, 0, 0, 0);
    
    const weekKey = weekStart.toISOString().split('T')[0];
    
    // Initialize week if not exists
    if (!weeklyData[weekKey]) {
      weeklyData[weekKey] = {
        date: weekStart,
        callsOffered: 0,
        callsAccepted: 0,
        totalTalkTime: 0,
        totalAcw: 0,
        onQueueMinutes: 0,
        totalMinutes: 0,
        days: 0
      };
    }
    
    // Accumulate data
    const week = weeklyData[weekKey];
    week.callsOffered += day.callsOffered || 0;
    week.callsAccepted += day.callsAccepted || 0;
    week.totalTalkTime += (day.averageTalkTime || 0) * (day.callsAccepted || 0);
    week.totalAcw += (day.acwAverage || 0) * (day.callsAccepted || 0);
    week.onQueueMinutes += (day.onQueuePercentage || 0) * 1440; // Minutes per day
    week.totalMinutes += 1440; // 24 hours in minutes
    week.days += 1;
  });
  
  // Calculate averages for each week
  Object.keys(weeklyData).forEach(weekKey => {
    const week = weeklyData[weekKey];
    
    // Calculate derived metrics
    week.answerRate = week.callsOffered ? week.callsAccepted / week.callsOffered : 0;
    week.averageTalkTime = week.callsAccepted ? week.totalTalkTime / week.callsAccepted : 0;
    week.acwAverage = week.callsAccepted ? week.totalAcw / week.callsAccepted : 0;
    week.onQueuePercentage = week.totalMinutes ? week.onQueueMinutes / week.totalMinutes : 0;
  });
  
  return weeklyData;
}

/**
 * Aggregate metrics by month
 * @param {Array} dailyMetrics - Daily metrics data
 * @return {Object} Monthly aggregated data
 */
function aggregateMetricsByMonth(dailyMetrics) {
  if (!dailyMetrics || !Array.isArray(dailyMetrics)) return {};
  
  const monthlyData = {};
  
  // Group by month
  dailyMetrics.forEach(day => {
    const date = new Date(day.date);
    
    // Get start of month
    const monthStart = new Date(date.getFullYear(), date.getMonth(), 1);
    
    const monthKey = monthStart.toISOString().split('T')[0];
    
    // Initialize month if not exists
    if (!monthlyData[monthKey]) {
      monthlyData[monthKey] = {
        date: monthStart,
        callsOffered: 0,
        callsAccepted: 0,
        totalTalkTime: 0,
        totalAcw: 0,
        onQueueMinutes: 0,
        totalMinutes: 0,
        days: 0
      };
    }
    
    // Accumulate data
    const month = monthlyData[monthKey];
    month.callsOffered += day.callsOffered || 0;
    month.callsAccepted += day.callsAccepted || 0;
    month.totalTalkTime += (day.averageTalkTime || 0) * (day.callsAccepted || 0);
    month.totalAcw += (day.acwAverage || 0) * (day.callsAccepted || 0);
    month.onQueueMinutes += (day.onQueuePercentage || 0) * 1440; // Minutes per day
    month.totalMinutes += 1440; // 24 hours in minutes
    month.days += 1;
  });
  
  // Calculate averages for each month
  Object.keys(monthlyData).forEach(monthKey => {
    const month = monthlyData[monthKey];
    
    // Calculate derived metrics
    month.answerRate = month.callsOffered ? month.callsAccepted / month.callsOffered : 0;
    month.averageTalkTime = month.callsAccepted ? month.totalTalkTime / month.callsAccepted : 0;
    month.acwAverage = month.callsAccepted ? month.totalAcw / month.callsAccepted : 0;
    month.onQueuePercentage = month.totalMinutes ? month.onQueueMinutes / month.totalMinutes : 0;
  });
  
  return monthlyData;
}

/**
 * Calculate team averages from individual metrics
 * @param {Array} teamMembers - Team members with metrics
 * @return {Object} Team average metrics
 */
function calculateTeamAverages(teamMembers) {
  if (!teamMembers || !Array.isArray(teamMembers) || teamMembers.length === 0) {
    return {
      answerRate: 0,
      averageTalkTime: 0,
      acwAverage: 0,
      onQueuePercentage: 0,
      callsAccepted: 0,
      score: 0
    };
  }
  
  // Initialize totals
  const totals = {
    callsAccepted: 0,
    weightedAnswerRate: 0,
    weightedTalkTime: 0,
    weightedAcw: 0,
    weightedOnQueue: 0,
    totalScore: 0
  };
  
  // Calculate weighted totals
  teamMembers.forEach(member => {
    const calls = member.callsAccepted || 0;
    totals.callsAccepted += calls;
    totals.weightedAnswerRate += (member.answerRate || 0) * calls;
    totals.weightedTalkTime += (member.averageTalkTime || 0) * calls;
    totals.weightedAcw += (member.acwAverage || 0) * calls;
    totals.weightedOnQueue += (member.onQueuePercentage || 0) * calls;
    totals.totalScore += member.score || 0;
  });
  
  // Calculate averages
  const averages = {
    answerRate: totals.callsAccepted ? totals.weightedAnswerRate / totals.callsAccepted : 0,
    averageTalkTime: totals.callsAccepted ? totals.weightedTalkTime / totals.callsAccepted : 0,
    acwAverage: totals.callsAccepted ? totals.weightedAcw / totals.callsAccepted : 0,
    onQueuePercentage: totals.callsAccepted ? totals.weightedOnQueue / totals.callsAccepted : 0,
    callsAccepted: Math.round(totals.callsAccepted / teamMembers.length), // Average calls per team member
    score: totals.totalScore / teamMembers.length
  };
  
  return averages;
}

/**
 * Apply formatting to a metric cell based on field type
 * @param {Sheet} sheet - Sheet object
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} field - Field name
 * @param {any} value - Metric value
 */
function applyMetricFormatting(sheet, row, col, field, value) {
  const cell = sheet.getRange(row, col);
  
  // Apply number format based on field type
  switch (field) {
    case 'Answer Rate':
    case 'On Queue Time':
    case 'Off Queue Time':
    case 'Interacting Time':
      cell.setNumberFormat("0.0%");
      break;
    case 'Talk Time':
    case 'After Call Work':
      cell.setNumberFormat("0.0 \"min\"");
      break;
    case 'Calls Offered':
    case 'Calls Accepted':
    case 'Staff Count':
      cell.setNumberFormat("#,##0");
      break;
    default:
      // Default number format if value is numeric
      if (typeof value === 'number') {
        if (value % 1 === 0) {
          cell.setNumberFormat("#,##0");
        } else {
          cell.setNumberFormat("0.00");
        }
      }
  }
  
  // Apply conditional formatting for key metrics
  if (field === 'Answer Rate') {
    // Answer Rate: ≥95% is good (green), ≥90% is average (yellow), <90% is poor (red)
    if (value >= 0.95) {
      cell.setBackground('#E6F4EA').setFontColor('#34A853');
    } else if (value >= 0.9) {
      cell.setBackground('#FEF7E0').setFontColor('#FBBC05');
    } else if (value > 0) { // Only apply if value is not zero (which might be a placeholder)
      cell.setBackground('#FDEDED').setFontColor('#EA4335');
    }
  } else if (field === 'Talk Time') {
    // Talk Time: 15-20 min is good (green), 13-15 or 20-22 is average (yellow), <13 or >22 is poor (red)
    if (value >= 15 && value <= 20) {
      cell.setBackground('#E6F4EA').setFontColor('#34A853');
    } else if ((value >= 13 && value < 15) || (value > 20 && value <= 22)) {
      cell.setBackground('#FEF7E0').setFontColor('#FBBC05');
    } else if (value > 0) { // Only apply if value is not zero (which might be a placeholder)
      cell.setBackground('#FDEDED').setFontColor('#EA4335');
    }
  } else if (field === 'After Call Work') {
    // ACW: ≤5 min is good (green), 5-7 min is average (yellow), >7 min is poor (red)
    if (value <= 5) {
      cell.setBackground('#E6F4EA').setFontColor('#34A853');
    } else if (value <= 7) {
      cell.setBackground('#FEF7E0').setFontColor('#FBBC05');
    } else if (value > 0) { // Only apply if value is not zero (which might be a placeholder)
      cell.setBackground('#FDEDED').setFontColor('#EA4335');
    }
  } else if (field === 'On Queue Time') {
    // On Queue: ≥65% is good (green), 60-65% is average (yellow), <60% is poor (red)
    if (value >= 0.65) {
      cell.setBackground('#E6F4EA').setFontColor('#34A853');
    } else if (value >= 0.6) {
      cell.setBackground('#FEF7E0').setFontColor('#FBBC05');
    } else if (value > 0) { // Only apply if value is not zero (which might be a placeholder)
      cell.setBackground('#FDEDED').setFontColor('#EA4335');
    }
  }
}

/**
 * Apply formatting to a column based on field type
 * @param {Sheet} sheet - Sheet object
 * @param {number} startRow - Starting row
 * @param {number} col - Column number
 * @param {number} numRows - Number of rows
 * @param {number} numCols - Number of columns
 * @param {string} field - Field name
 */
function applyColumnFormatting(sheet, startRow, col, numRows, numCols, field) {
  const range = sheet.getRange(startRow, col, numRows, numCols);
  
  // Apply number format based on field type
  switch (field) {
    case 'Answer Rate':
    case 'On Queue Time':
    case 'Off Queue Time':
    case 'Interacting Time':
      range.setNumberFormat("0.0%");
      break;
    case 'Talk Time':
    case 'After Call Work':
      range.setNumberFormat("0.0 \"min\"");
      break;
    case 'Calls Offered':
    case 'Calls Accepted':
    case 'Staff Count':
      range.setNumberFormat("#,##0");
      break;
    default:
      // Try to determine format from first non-null cell
      for (let row = startRow; row < startRow + numRows; row++) {
        const value = sheet.getRange(row, col).getValue();
        if (value !== null && value !== undefined) {
          if (typeof value === 'number') {
            if (value % 1 === 0) {
              range.setNumberFormat("#,##0");
            } else {
              range.setNumberFormat("0.00");
            }
          }
          break;
        }
      }
  }
  
  // Apply conditional formatting for key metrics
  if (field === 'Answer Rate') {
    range.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.95)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.9)
        .whenNumberLessThan(0.95)
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .whenNumberLessThan(0.9)
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([range])
        .build()
    ]);
  } else if (field === 'Talk Time') {
    range.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(15, 20)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=OR(AND(C2>=13,C2<15),AND(C2>20,C2<=22))')
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(C2>0,OR(C2<13,C2>22))')
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([range])
        .build()
    ]);
  } else if (field === 'After Call Work') {
    range.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThanOrEqualTo(5)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(5)
        .whenNumberLessThanOrEqualTo(7)
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(7)
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([range])
        .build()
    ]);
  } else if (field === 'On Queue Time') {
    range.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.65)
        .setBackground('#E6F4EA')
        .setFontColor('#34A853')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.6)
        .whenNumberLessThan(0.65)
        .setBackground('#FEF7E0')
        .setFontColor('#FBBC05')
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .whenNumberLessThan(0.6)
        .setBackground('#FDEDED')
        .setFontColor('#EA4335')
        .setRanges([range])
        .build()
    ]);
  }
}

/**
 * Format month for display
 * @param {Date} date - Date object
 * @return {string} Formatted month string
 */
function formatMonthForDisplay(date) {
  if (!date) return '';
  
  try {
    return date.toLocaleDateString(undefined, { year: 'numeric', month: 'long' });
  } catch (error) {
    return `${date.getMonth() + 1}/${date.getFullYear()}`;
  }
}

/**
 * Add charts to spreadsheet
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {Object} config - Export configuration
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function addChartsToSpreadsheet(spreadsheet, config, startDate, endDate) {
  try {
    // Only add charts if requested
    if (!config.includeCharts) return;
    
    // Create charts sheet
    const chartsSheet = spreadsheet.insertSheet("Charts");
    
    // Get data sheets based on export type
    let dataSheetNames = [];
    if (config.exportType === 'team' || config.exportType === 'comprehensive') {
      // Look for data sheets
      const sheets = spreadsheet.getSheets();
      for (const sheet of sheets) {
        const sheetName = sheet.getName();
        if (sheetName === "Daily Metrics" || sheetName === "Weekly Metrics" || sheetName === "Monthly Metrics") {
          dataSheetNames.push(sheetName);
        }
      }
    }
    
    if (dataSheetNames.length === 0) {
      chartsSheet.getRange(1, 1).setValue("No data available for charts");
      return;
    }
    
    // Create charts based on available data
    let currentRow = 1;
    
    // Use daily metrics if available
    if (dataSheetNames.includes("Daily Metrics")) {
      const dataSheet = spreadsheet.getSheetByName("Daily Metrics");
      const dataRange = dataSheet.getDataRange();
      const lastRow = dataRange.getLastRow();
      const numRows = lastRow - 1; // Exclude header row
      
      if (numRows > 0) {
        // Find column numbers for key metrics
        const headers = dataSheet.getRange(1, 1, 1, dataRange.getLastColumn()).getValues()[0];
        const metrics = {
          answerRate: headers.indexOf('Answer Rate') + 1,
          callsAccepted: headers.indexOf('Calls Accepted') + 1,
          callsOffered: headers.indexOf('Calls Offered') + 1,
          talkTime: headers.indexOf('Talk Time') + 1,
          onQueueTime: headers.indexOf('On Queue Time') + 1
        };
        
        // Call volume chart
        if (metrics.callsOffered > 0) {
          chartsSheet.getRange(currentRow, 1).setValue("Call Volume Trend");
          chartsSheet.getRange(currentRow, 1).setFontWeight("bold");
          currentRow++;
          
          const callVolumeChart = chartsSheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(dataSheet.getRange(1, 1, numRows + 1, 1)) // Date column
            .addRange(dataSheet.getRange(1, metrics.callsOffered, numRows + 1, 1)) // Calls offered
            .setPosition(currentRow, 1, 0, 0)
            .setOption('title', 'Call Volume Trend')
            .setOption('legend', {position: 'top'})
            .setOption('width', 600)
            .setOption('height', 300)
            .setOption('hAxis', {title: 'Date', textStyle: {fontSize: 10}})
            .setOption('vAxis', {title: 'Calls'})
            .setOption('colors', ['#4285F4'])
            .build();
          
          chartsSheet.insertChart(callVolumeChart);
          currentRow += 16; // Space for chart
        }
        
        // Answer rate chart
        if (metrics.answerRate > 0) {
          chartsSheet.getRange(currentRow, 1).setValue("Answer Rate Trend");
          chartsSheet.getRange(currentRow, 1).setFontWeight("bold");
          currentRow++;
          
          const answerRateChart = chartsSheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(dataSheet.getRange(1, 1, numRows + 1, 1)) // Date column
            .addRange(dataSheet.getRange(1, metrics.answerRate, numRows + 1, 1)) // Answer rate
            .setPosition(currentRow, 1, 0, 0)
            .setOption('title', 'Answer Rate Trend')
            .setOption('legend', {position: 'top'})
            .setOption('width', 600)
            .setOption('height', 300)
            .setOption('hAxis', {title: 'Date', textStyle: {fontSize: 10}})
            .setOption('vAxis', {
              title: 'Answer Rate', 
              format: 'percent',
              viewWindow: {min: 0.8, max: 1}
            })
            .setOption('colors', ['#34A853'])
            .build();
          
          chartsSheet.insertChart(answerRateChart);
          currentRow += 16; // Space for chart
        }
        
        // Create a day-of-week analysis
        if (metrics.callsOffered > 0 && numRows >= 7) {
          chartsSheet.getRange(currentRow, 1).setValue("Call Volume by Day of Week");
          chartsSheet.getRange(currentRow, 1).setFontWeight("bold");
          currentRow++;
          
          // Create aggregation by day of week
          chartsSheet.getRange(currentRow, 1).setValue("Day of Week");
          chartsSheet.getRange(currentRow, 2).setValue("Average Calls");
          chartsSheet.getRange(currentRow, 1, 1, 2).setFontWeight("bold");
          currentRow++;
          
          // Initialize totals and counts for each day of week
          const dayOfWeekData = {};
          for (let i = 0; i < 7; i++) {
            dayOfWeekData[i] = { total: 0, count: 0 };
          }
          
          // Aggregate data
          for (let row = 2; row <= lastRow; row++) {
            const date = dataSheet.getRange(row, 1).getValue();
            if (date instanceof Date) {
              const dayOfWeek = date.getDay(); // 0=Sunday, 6=Saturday
              const callsOffered = dataSheet.getRange(row, metrics.callsOffered).getValue() || 0;
              
              dayOfWeekData[dayOfWeek].total += callsOffered;
              dayOfWeekData[dayOfWeek].count++;
            }
          }
          
          // Write aggregated data
          const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          const dowData = [];
          
          for (let i = 0; i < 7; i++) {
            const avg = dayOfWeekData[i].count ? dayOfWeekData[i].total / dayOfWeekData[i].count : 0;
            dowData.push([dayNames[i], avg]);
            chartsSheet.getRange(currentRow, 1).setValue(dayNames[i]);
            chartsSheet.getRange(currentRow, 2).setValue(avg);
            currentRow++;
          }
          
          // Add day of week chart
          const dowChart = chartsSheet.newChart()
            .setChartType(Charts.ChartType.COLUMN)
            .addRange(chartsSheet.getRange(currentRow - 7, 1, 7, 2)) // Day of week data
            .setPosition(currentRow, 1, 0, 0)
            .setOption('title', 'Average Call Volume by Day of Week')
            .setOption('legend', {position: 'none'})
            .setOption('width', 600)
            .setOption('height', 300)
            .setOption('hAxis', {title: 'Day of Week'})
            .setOption('vAxis', {title: 'Average Calls'})
            .setOption('colors', ['#4285F4'])
            .build();
          
          chartsSheet.insertChart(dowChart);
          currentRow += 16; // Space for chart
        }
      }
    }
    
    // Auto-resize columns
    chartsSheet.autoResizeColumns(1, 10);
  } catch (error) {
    console.error("Error adding charts to spreadsheet: " + error.message);
    // Add error message to charts sheet
    try {
      const chartsSheet = spreadsheet.getSheetByName("Charts");
      if (chartsSheet) {
        chartsSheet.getRange(1, 1).setValue("Error generating charts: " + error.message);
      }
    } catch (e) {
      // Ignore any errors in error handling
    }
  }
}

/**
 * Generate chart data for email templates
 * @param {Object} metricsData - Metrics data
 * @param {string} metricKey - Metric key to graph
 * @param {number} [multiplier=1] - Value multiplier (e.g., 100 for percentages)
 * @return {string} Comma-separated data values for Google Charts
 */
function generateChartData(metricsData, metricKey, multiplier = 1) {
  if (!metricsData || !metricsData.dailyMetrics || metricsData.dailyMetrics.length === 0) {
    return '0';
  }
  
  // Get values for the selected metric
  const values = metricsData.dailyMetrics.map(day => {
    const rawValue = getMetricValueByName(metricKey, day);
    return (rawValue || 0) * multiplier;
  });
  
  return values.join(',');
}

/**
 * Generate chart labels for email templates
 * @param {Object} metricsData - Metrics data
 * @return {string} Pipe-separated date labels for Google Charts
 */
function generateChartLabels(metricsData) {
  if (!metricsData || !metricsData.dailyMetrics || metricsData.dailyMetrics.length === 0) {
    return '';
  }
  
  // Format dates for short display
  const formatDate = date => {
    try {
      return new Date(date).toLocaleDateString(undefined, { month: 'numeric', day: 'numeric' });
    } catch (e) {
      return date;
    }
  };
  
  // Get dates
  let labels = metricsData.dailyMetrics.map(day => formatDate(day.date));
  
  // For readability, if we have more than 10 days, only show every other/third label
  if (labels.length > 14) {
    labels = labels.map((label, i) => i % 3 === 0 ? label : ' ');
  } else if (labels.length > 7) {
    labels = labels.map((label, i) => i % 2 === 0 ? label : ' ');
  }
  
  return labels.join('|');
}

/**
 * Generate scheduled export
 * Called by time-based trigger
 */
function generateScheduledExport() {
  try {
    // Load export config
    const exportConfig = loadExportConfig();
    
    // Skip if scheduling is disabled
    if (!exportConfig.scheduleExport) {
      console.log("Export scheduling is disabled.");
      return;
    }
    
    // Generate export
    const result = generateExportNow(exportConfig);
    
    if (result.success) {
      // Log success
      console.log(`Scheduled export generated successfully: ${result.fileName}`);
      
      // Send email notification if configured
      const userEmail = UserService.getCurrentUserEmail();
      MailApp.sendEmail({
        to: userEmail,
        subject: "988 Crisis Line Metrics - Scheduled Export Ready",
        htmlBody: `
          <p>Your scheduled metrics export has been generated.</p>
          <p><strong>File:</strong> ${result.fileName}</p>
          <p>You can access this file in your Google Drive under "988 Crisis Line Exports".</p>
          <p>This is an automated message from the 988 Crisis Line Metrics System.</p>
        `
      });
    } else {
      // Log error
      console.error(`Failed to generate scheduled export: ${result.message}`);
    }
  } catch (error) {
    console.error("Error in scheduled export: " + error.message);
    
    // Send error notification
    try {
      const userEmail = UserService.getCurrentUserEmail();
      MailApp.sendEmail({
        to: userEmail,
        subject: "Error: 988 Crisis Line Scheduled Export Failed",
        body: `There was an error generating the scheduled export:\n\n${error.message}\n\nPlease check the configuration.`
      });
    } catch (e) {
      console.error("Failed to send error notification: " + e.message);
    }
  }
}

/**
 * Send test email to verify configuration
 * @param {Object} [config] - Optional email config (uses stored config if not provided)
 * @return {Object} Result with success flag and message
 */
function sendTestEmail(config = null) {
  try {
    // Use provided config or load from storage
    const emailConfig = config || loadEmailConfig();
    
    // Get current user's email
    const userEmail = UserService.getCurrentUserEmail();
    
    // Generate test content with current date range
    const endDate = new Date();
    const startDate = getStartDateFromConfig(emailConfig.dateRange, endDate);
    
    // Generate email content
    const htmlContent = generateEmailContent(emailConfig, startDate, endDate);
    const textContent = generateTextEmailContent(emailConfig, startDate, endDate);
    
    // Format the subject for test email
    const subject = `TEST: ${formatEmailSubject(emailConfig.emailSubject, startDate, endDate)}`;
    
    // Send test email to current user
    const fromName = emailConfig.fromName || "988 Crisis Line Analytics";
    
    if (emailConfig.reportFormat === 'rich') {
      MailApp.sendEmail({
        to: userEmail,
        subject: subject,
        htmlBody: htmlContent,
        name: fromName,
        replyTo: emailConfig.replyToEmail || ''
      });
    } else {
      MailApp.sendEmail({
        to: userEmail,
        subject: subject,
        body: textContent,
        name: fromName,
        replyTo: emailConfig.replyToEmail || ''
      });
    }
    
    return {
      success: true,
      message: "Test email sent to your email address"
    };
  } catch (error) {
    console.error("Error sending test email: " + error.message);
    return {
      success: false,
      message: "Error sending test email: " + error.message
    };
  }
}

/**
 * Log export error with details
 * @param {Error} error - Error object
 * @param {Object} context - Error context
 */
function logExportError(error, context) {
  try {
    // Get information about the context
    const user = UserService.getCurrentUserEmail() || "unknown";
    const timestamp = new Date().toISOString();
    
    // Prepare log data
    const logData = {
      timestamp: timestamp,
      user: user,
      error: error.message,
      stack: error.stack,
      context: context
    };
    
    // Log to console
    console.error("Export Error:", logData);
    
    // Store in error log (optional)
    try {
      const errorLogs = PropertiesService.getScriptProperties().getProperty('metrics_export_errors') || '[]';
      const logs = JSON.parse(errorLogs);
      
      // Limit error logs to most recent 100
      logs.unshift(logData);
      if (logs.length > 100) logs.length = 100;
      
      PropertiesService.getScriptProperties().setProperty('metrics_export_errors', JSON.stringify(logs));
    } catch (e) {
      console.error("Error saving to error log: " + e.message);
    }
  } catch (e) {
    console.error("Error in error logging: " + e.message);
  }
}

/**
 * Get metrics overview for dashboard
 * @param {string} [period='week'] - Time period (day, week, month, quarter, year)
 * @return {Object} Metrics overview
 */
function getMetricsOverview(period = 'week') {
  try {
    // Calculate date range based on period
    const endDate = new Date();
    let startDate;
    
    switch (period) {
      case 'day':
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 1);
        break;
      case 'week':
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 7);
        break;
      case 'month':
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 30);
        break;
      case 'quarter':
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 90);
        break;
      case 'year':
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 365);
        break;
      default:
        startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - 7);
    }
    
    // Get metrics data
    const metricsData = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Also get data for previous period for comparison
    const previousEndDate = new Date(startDate);
    const periodDays = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24));
    const previousStartDate = new Date(previousEndDate);
    previousStartDate.setDate(previousStartDate.getDate() - periodDays);
    
    const previousData = MetricsService.getTeamMetrics(previousStartDate, previousEndDate);
    
    // Format the data for dashboard display with comparisons
    return {
      period: period,
      dateRange: {
        start: startDate.toISOString(),
        end: endDate.toISOString(),
        display: `${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)}`
      },
      metrics: {
        answerRate: {
          value: metricsData.answerRate,
          previous: previousData.answerRate,
          change: calculateChange(metricsData.answerRate, previousData.answerRate),
          target: 0.95,
          format: 'percent'
        },
        averageTalkTime: {
          value: metricsData.averageTalkTime,
          previous: previousData.averageTalkTime,
          change: calculateChange(metricsData.averageTalkTime, previousData.averageTalkTime),
          target: 17.5, // Target is 15-20 minutes, so midpoint
          format: 'time'
        },
        acwAverage: {
          value: metricsData.acwAverage,
          previous: previousData.acwAverage,
          change: calculateChange(metricsData.acwAverage, previousData.acwAverage, true), // lower is better
          target: 5,
          format: 'time'
        },
        onQueuePercentage: {
          value: metricsData.onQueuePercentage,
          previous: previousData.onQueuePercentage,
          change: calculateChange(metricsData.onQueuePercentage, previousData.onQueuePercentage),
          target: 0.65,
          format: 'percent'
        },
        callsOffered: {
          value: metricsData.callsOffered,
          previous: previousData.callsOffered,
          change: calculateChange(metricsData.callsOffered, previousData.callsOffered),
          format: 'number'
        },
        callsAccepted: {
          value: metricsData.callsAccepted,
          previous: previousData.callsAccepted,
          change: calculateChange(metricsData.callsAccepted, previousData.callsAccepted),
          format: 'number'
        }
      },
      teamPerformance: {
        averageScore: metricsData.teamScore || 0,
        topPerformer: metricsData.topPerformer || 'N/A',
        totalTeamMembers: metricsData.totalTeamMembers || 0
      },
      dailyTrend: metricsData.dailyMetrics || []
    };
  } catch (error) {
    console.error("Error getting metrics overview: " + error.message);
    
    // Return empty data structure with error flag
    return {
      error: true,
      message: error.message,
      period: period,
      dateRange: {
        start: new Date().toISOString(),
        end: new Date().toISOString(),
        display: 'Error loading data'
      },
      metrics: {},
      teamPerformance: {},
      dailyTrend: []
    };
  }
}

/**
 * Calculate change between current and previous values
 * @param {number} current - Current value
 * @param {number} previous - Previous value
 * @param {boolean} [lowerIsBetter=false] - Whether lower values are better
 * @return {Object} Change information
 */
function calculateChange(current, previous, lowerIsBetter = false) {
  if (current === null || current === undefined || previous === null || previous === undefined) {
    return { percent: 0, direction: 'none' };
  }
  
  if (previous === 0) {
    return { percent: 100, direction: lowerIsBetter ? 'negative' : 'positive' };
  }
  
  const percentChange = ((current - previous) / Math.abs(previous)) * 100;
  let direction = 'none';
  
  if (percentChange > 0) {
    direction = lowerIsBetter ? 'negative' : 'positive';
  } else if (percentChange < 0) {
    direction = lowerIsBetter ? 'positive' : 'negative';
  }
  
  return {
    percent: Math.abs(percentChange),
    direction: direction
  };
}

/**
 * Helper function to convert string to camelCase
 * @param {string} str - Input string
 * @return {string} camelCase string
 */
function toCamelCase(str) {
  return str
    .toLowerCase()
    .replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase());
}

/**
 * Helper function to convert string to snake_case
 * @param {string} str - Input string
 * @return {string} snake_case string
 */
function toSnakeCase(str) {
  return str
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^a-zA-Z0-9_]/g, '');
}

/**
 * Test email template generation
 * @return {Object} Test result with generated templates
 */
function testEmailTemplate() {
  try {
    // Load default config
    const config = DEFAULT_CONFIG.email;
    
    // Set test date range
    const endDate = new Date();
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 7);
    
    // Generate templates
    const richHtml = generateRichEmailContent(config, startDate, endDate);
    const simpleHtml = generateSimpleEmailContent(config, startDate, endDate);
    const text = generateTextEmailContent(config, startDate, endDate);
    
    return {
      success: true,
      rich: richHtml,
      simple: simpleHtml,
      text: text
    };
  } catch (error) {
    console.error("Error testing email template: " + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Test export generation without saving
 * @return {Object} Test result
 */
function testExportGeneration() {
  try {
    // Load default config
    const config = DEFAULT_CONFIG.export;
    
    // Set test date range
    const endDate = new Date();
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 7);
    
    // Get metrics data for test
    const metricsData = MetricsService.getTeamMetrics(startDate, endDate);
    
    // Test CSV generation
    const csvContent = generateTeamMetricsCsv({
      ...config,
      fields: ['Answer Rate', 'Talk Time', 'After Call Work']
    }, startDate, endDate);
    
    // For spreadsheet and PDF, just check if we could generate them
    // but don't actually create the files to save resources
    
    return {
      success: true,
      metricsAvailable: !!metricsData,
      csvSample: csvContent.substring(0, 500) + '...' // Just return a sample
    };
  } catch (error) {
    console.error("Error testing export generation: " + error.message);
    return {
      success: false,
      message: error.message
    };
  }
}  
