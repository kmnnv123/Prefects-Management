// ===================================
// CALENDAR MODULE - BOP25 Attendance System
// Contains all calendar-related functions
// ===================================

console.log('✅ Calendar module loaded successfully');

// Format date for comparison (YYYY-MM-DD) - Fixed timezone handling
function formatDateForComparison(date) {
    // Use local timezone to avoid date shifting
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// Calculate attendance levels for color coding
function calculateAttendanceLevels(availableDates) {
    const levels = {};
    
    availableDates.forEach(dateStr => {
        const attendedCount = extractedData.filter(employee => 
            employee.attendanceData.some(record => 
                record.fullDate && 
                formatDateForComparison(record.fullDate) === dateStr &&
                record.morning.in
            )
        ).length;
        
        // Color coding based on attendance numbers:
        // Green: above 45 attendees
        // Yellow: 30-45 attendees  
        // Red: below 30 attendees
        if (attendedCount > 45) {
            levels[dateStr] = 'high';
        } else if (attendedCount >= 30) {
            levels[dateStr] = 'moderate';
        } else {
            levels[dateStr] = 'low';
        }
    });
    
    return levels;
}

// Generate year options for year selector
function generateYearOptions(currentYear) {
    const startYear = currentYear - 5;
    const endYear = currentYear + 5;
    let options = '';
    
    for (let year = startYear; year <= endYear; year++) {
        options += `<option value="${year}" ${year === currentYear ? 'selected' : ''}>${year}</option>`;
    }
    
    return options;
}

// Change calendar month (navigation buttons)
function changeCalendarMonth(direction) {
    if (!window.currentCalendarDate) return;
    
    window.currentCalendarDate.setMonth(window.currentCalendarDate.getMonth() + direction);
    window.renderCalendar();
}

// Change calendar to specific month
function changeCalendarToMonth(monthIndex) {
    if (!window.currentCalendarDate) return;
    
    window.currentCalendarDate.setMonth(parseInt(monthIndex));
    window.renderCalendar();
}

// Change calendar to specific year
function changeCalendarToYear(year) {
    if (!window.currentCalendarDate) return;
    
    window.currentCalendarDate.setFullYear(parseInt(year));
    window.renderCalendar();
}

// Initialize the attendance calendar
function initializeAttendanceCalendar(availableDates, resultsDivId) {
    const calendarContainer = document.getElementById('attendanceCalendar');
    
    if (!calendarContainer) {
        console.error('❌ Calendar container not found!');
        return;
    }
    
    const today = new Date();
    let currentDate = new Date(today.getFullYear(), today.getMonth(), 1);
    
    // Convert available dates to a Set for quick lookup
    const dateSet = new Set(availableDates);
    
    // Calculate attendance levels for each date
    const attendanceLevels = calculateAttendanceLevels(availableDates);

    function renderCalendar() {
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();
        
        // Get first day of month and number of days
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        const daysInMonth = lastDay.getDate();
        const startingDayOfWeek = firstDay.getDay();
        
        // Get previous month's last days to fill the grid
        const prevMonth = new Date(year, month, 0);
        const daysInPrevMonth = prevMonth.getDate();
        
        const monthNames = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ];
        
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        
        let calendarHtml = `
            <div class="calendar-header">
                <div class="calendar-navigation">
                    <button onclick="changeCalendarMonth(-1)" class="btn btn-sm btn-outline-success">
                        <i class="bi bi-chevron-left"></i>
                    </button>
                    <div class="calendar-month-year">
                        <select onchange="changeCalendarToMonth(this.value)" class="month-selector">
                            ${monthNames.map((name, index) => 
                                `<option value="${index}" ${index === month ? 'selected' : ''}>${name}</option>`
                            ).join('')}
                        </select>
                        <select onchange="changeCalendarToYear(this.value)" class="year-selector">
                            ${generateYearOptions(year)}
                        </select>
                    </div>
                    <button onclick="changeCalendarMonth(1)" class="btn btn-sm btn-outline-success">
                        <i class="bi bi-chevron-right"></i>
                    </button>
                </div>
            </div>
            
            <div class="calendar-grid">
        `;
        
        // Day headers
        dayNames.forEach(day => {
            calendarHtml += `<div class="calendar-day-header">${day}</div>`;
        });
        
        // Previous month's days (grayed out)
        for (let i = startingDayOfWeek - 1; i >= 0; i--) {
            const day = daysInPrevMonth - i;
            const date = new Date(year, month - 1, day);
            const dateStr = formatDateForComparison(date);
            const hasData = dateSet.has(dateStr);
            const attendanceLevel = attendanceLevels[dateStr] || 'none';
            const isHolidayDate = isHoliday(dateStr);
            const isFutureDate = date > today;
            
            calendarHtml += `
                <div class="calendar-day other-month ${hasData ? 'has-data' : ''} ${isHolidayDate ? 'holiday' : ''} ${isFutureDate ? 'future-date' : ''} attendance-${attendanceLevel}" 
                     onclick="selectCalendarDate('${dateStr}', '${resultsDivId}')">
                    ${day}
                </div>
            `;
        }
        
        // Current month's days
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(year, month, day);
            const dateStr = formatDateForComparison(date);
            const hasData = dateSet.has(dateStr);
            const isToday = date.toDateString() === today.toDateString();
            const attendanceLevel = attendanceLevels[dateStr] || 'none';
            const isHolidayDate = isHoliday(dateStr);
            const isFutureDate = date > today;
            
            let classes = 'calendar-day';
            if (isToday) classes += ' today';
            if (hasData) classes += ' has-data';
            if (isHolidayDate) classes += ' holiday';
            if (isFutureDate) classes += ' future-date';
            classes += ` attendance-${attendanceLevel}`;
            
            calendarHtml += `
                <div class="${classes}" onclick="selectCalendarDate('${dateStr}', '${resultsDivId}')">
                    ${day}
                </div>
            `;
        }
        
        // Next month's days to fill the grid (6 rows total)
        const totalCells = Math.ceil((startingDayOfWeek + daysInMonth) / 7) * 7;
        const remainingCells = totalCells - (startingDayOfWeek + daysInMonth);
        
        for (let day = 1; day <= remainingCells; day++) {
            const date = new Date(year, month + 1, day);
            const dateStr = formatDateForComparison(date);
            const hasData = dateSet.has(dateStr);
            const attendanceLevel = attendanceLevels[dateStr] || 'none';
            const isHolidayDate = isHoliday(dateStr);
            const isFutureDate = date > today;
            
            calendarHtml += `
                <div class="calendar-day other-month ${hasData ? 'has-data' : ''} ${isHolidayDate ? 'holiday' : ''} ${isFutureDate ? 'future-date' : ''} attendance-${attendanceLevel}" 
                     onclick="selectCalendarDate('${dateStr}', '${resultsDivId}')">
                    ${day}
                </div>
            `;
        }
        
        calendarHtml += `
            </div>
        `;
        
        calendarContainer.innerHTML = calendarHtml;
        
        // Add smooth fade-in animation for calendar days
        requestAnimationFrame(() => {
            const calendarDays = calendarContainer.querySelectorAll('.calendar-day');
            calendarDays.forEach((day, index) => {
                day.style.opacity = '0';
                day.style.transform = 'translateY(10px)';
                day.style.transition = `all 0.3s cubic-bezier(0.4, 0, 0.2, 1)`;
                
                // Stagger the animation for each day
                setTimeout(() => {
                    day.style.opacity = '1';
                    day.style.transform = 'translateY(0)';
                }, index * 8); // Small delay between each day
            });
        });
    }
    
    // Store calendar state globally for navigation
    window.currentCalendarDate = currentDate;
    window.calendarAvailableDates = dateSet;
    window.calendarResultsDiv = resultsDivId;
    window.calendarAttendanceLevels = attendanceLevels;
    window.renderCalendar = renderCalendar;
    
    renderCalendar();
}

// Select calendar date with timezone-safe handling
function selectCalendarDate(dateStr, resultsDivId) {
    // Check if clicked date is a holiday
    if (isHoliday(dateStr)) {
        // Show holiday message
        alert('This date is marked as a holiday. Attendance data is not tracked for holidays.');
        return;
    }

    // Clear previous selection and reset inline styles
    document.querySelectorAll('.calendar-day.selected').forEach(day => {
        day.classList.remove('selected');
        day.style.background = '';
        day.style.borderColor = '';
        day.style.color = '';
    });
    
    // Mark new selection with green highlighting
    const clickedElement = event.target;
    clickedElement.classList.add('selected');
    
    console.log(`📅 Calendar date selected: ${dateStr}`);
    
    // Load attendance for the selected date
    loadDailyAttendanceForDate(dateStr, resultsDivId);
    
    // Auto-collapse calendar after selection with visual feedback
    const calendar = document.getElementById('attendanceCalendar');
    
    if (calendar.classList.contains('calendar-expanded')) {
        // Add selection feedback animation
        clickedElement.style.transition = 'all 0.2s ease';
        clickedElement.style.transform = 'scale(0.95)';
        
        setTimeout(() => {
            clickedElement.style.transform = 'scale(1)';
        }, 100);
        
        // Smooth collapse after selection
        setTimeout(() => {
            collapseCalendar();
        }, 600); // Increased delay for better UX
    }
}

// Toggle calendar visibility (mobile and desktop)
function toggleCalendarView() {
    const calendar = document.getElementById('attendanceCalendar');
    const toggleBtn = document.getElementById('calendarToggleBtn');
    const toggleIcon = toggleBtn.querySelector('i');
    
    if (calendar.classList.contains('calendar-collapsed')) {
        // Show calendar with smooth animation
        calendar.style.transition = 'all 0.4s cubic-bezier(0.4, 0, 0.2, 1)';
        calendar.style.transform = 'translateY(-10px)';
        calendar.style.opacity = '0';
        
        calendar.classList.remove('calendar-collapsed');
        calendar.classList.add('calendar-expanded');
        
        // Animate the toggle button icon
        toggleIcon.style.transition = 'transform 0.3s ease';
        toggleIcon.style.transform = 'rotate(180deg)';
        
        // Start the smooth reveal animation
        requestAnimationFrame(() => {
            calendar.style.transform = 'translateY(0)';
            calendar.style.opacity = '1';
        });
        
        // Update button state after a short delay
        setTimeout(() => {
            toggleIcon.className = 'bi bi-x-lg';
            toggleIcon.style.transform = 'rotate(0deg)';
            toggleBtn.classList.add('active');
            toggleBtn.title = 'Close Calendar';
        }, 150);
        
        // Add enhanced shadow effect
        setTimeout(() => {
            calendar.style.boxShadow = '0 12px 40px rgba(16, 185, 129, 0.2), 0 4px 16px rgba(16, 185, 129, 0.1)';
        }, 300);
    } else {
        // Hide calendar with smooth animation
        collapseCalendar();
    }
}

// Collapse calendar function (reusable)
function collapseCalendar() {
    const calendar = document.getElementById('attendanceCalendar');
    const toggleBtn = document.getElementById('calendarToggleBtn');
    const toggleIcon = toggleBtn.querySelector('i');
    
    // Start smooth collapse animation
    calendar.style.transition = 'all 0.4s cubic-bezier(0.4, 0, 0.2, 1)';
    calendar.style.transform = 'translateY(-10px)';
    calendar.style.opacity = '0';
    calendar.style.boxShadow = 'none';
    
    // Animate the toggle button icon
    toggleIcon.style.transition = 'transform 0.3s ease';
    toggleIcon.style.transform = 'rotate(-180deg)';
    
    // Update button state
    setTimeout(() => {
        toggleIcon.className = 'bi bi-calendar';
        toggleIcon.style.transform = 'rotate(0deg)';
        toggleBtn.classList.remove('active');
        toggleBtn.title = 'Toggle Calendar';
    }, 150);
    
    // Complete the collapse after animation
    setTimeout(() => {
        calendar.classList.remove('calendar-expanded');
        calendar.classList.add('calendar-collapsed');
        calendar.style.transform = '';
        calendar.style.opacity = '';
        calendar.style.transition = '';
    }, 400);
}

// Load today's attendance on initial load
function loadTodaysAttendance(resultsDivId) {
    const today = new Date();
    const todayStr = formatDateForComparison(today);
    
    // Check if today has data
    const hasDataForToday = extractedData.some(employee => 
        employee.attendanceData.some(record => 
            record.fullDate && formatDateForComparison(record.fullDate) === todayStr
        )
    );
    
    if (hasDataForToday) {
        console.log('📅 Loading today\'s attendance:', todayStr);
        loadDailyAttendanceForDate(todayStr, resultsDivId);
    } else {
        // Show today's placeholder in daily attendance area
        showNoDataMessage(todayStr, 'today');
    }
}

// Show no data message
function showNoDataMessage(dateStr, type = 'selected') {
    const today = new Date(dateStr + 'T00:00:00');
    const displayDate = today.toLocaleDateString('en-US', { 
        weekday: 'long', 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
    });
    
    const downloadBtn = document.getElementById('downloadDailyBtn');
    
    // Create or update the daily attendance display area
    let dailyDisplayArea = document.getElementById('dailyAttendanceDisplay');
    if (!dailyDisplayArea) {
        dailyDisplayArea = document.createElement('div');
        dailyDisplayArea.id = 'dailyAttendanceDisplay';
        dailyDisplayArea.style.marginTop = '2rem';
        
        // Insert after the controls but before any other content
        const dailyControls = document.getElementById('dailyAttendanceControls');
        dailyControls.parentNode.insertBefore(dailyDisplayArea, dailyControls.nextSibling);
    }
    
    const messageType = type === 'today' ? 'Today\'s' : 'Selected Date';
    const messageText = type === 'today' ? 
        'Attendance data for today has not been uploaded yet.' : 
        'No attendance data available for this date.';
    
    dailyDisplayArea.innerHTML = `
        <div class="no-data-container">
            <div class="no-data-message">
                <i class="bi bi-calendar-x" style="font-size: 3rem; opacity: 0.3; margin-bottom: 1rem;"></i>
                <h5 class="mt-3">${messageType} Data Not Available</h5>
                <p class="mb-2"><strong>${displayDate}</strong></p>
                <p class="text-muted">${messageText}</p>
                <small class="text-muted">
                    <i class="bi bi-info-circle me-1"></i>Click the calendar icon to view available dates
                </small>
            </div>
        </div>
        
        <style>
            .no-data-container {
                background: #1a1a1a;
                border: 1px solid #2d2d2d;
                border-radius: 16px;
                padding: 3rem 2rem;
                text-align: center;
                color: #6b7280;
                margin-top: 1rem;
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
            }
            
            .no-data-message h5 {
                color: #9ca3af;
                font-weight: 500;
            }
            
            .no-data-message p {
                margin-bottom: 0.5rem;
            }
            
            .no-data-message small {
                background: #374151;
                padding: 0.5rem 1rem;
                border-radius: 8px;
                display: inline-block;
                margin-top: 1rem;
            }
        </style>
    `;
    
    if (downloadBtn) {
        downloadBtn.disabled = true;
    }
}

// Holiday Calendar Functions
// Initialize the holiday calendar using the same system as daily attendance
function initializeHolidayCalendar(year, month) {
    const calendarContainer = document.getElementById('holidayCalendarContainer');
    const today = new Date();
    let currentDate = new Date(year, month - 1, 1);
    
    function renderHolidayCalendar() {
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();
        
        // Get first day of month and number of days
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        const daysInMonth = lastDay.getDate();
        const startingDayOfWeek = firstDay.getDay();
        
        // Get previous month's last days to fill the grid
        const prevMonth = new Date(year, month, 0);
        const daysInPrevMonth = prevMonth.getDate();
        
        const monthNames = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ];
        
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        
        let calendarHtml = `
            <div class="calendar-header">
                <div class="calendar-navigation">
                    <button onclick="changeHolidayCalendarMonth(-1)" class="btn btn-sm btn-outline-success">
                        <i class="bi bi-chevron-left"></i>
                    </button>
                    <div class="calendar-month-year">
                        <select onchange="changeHolidayCalendarToMonth(this.value)" class="month-selector">
                            ${monthNames.map((name, index) => 
                                `<option value="${index}" ${index === month ? 'selected' : ''}>${name}</option>`
                            ).join('')}
                        </select>
                        <select onchange="changeHolidayCalendarToYear(this.value)" class="year-selector">
                            ${generateYearOptions(year)}
                        </select>
                    </div>
                    <button onclick="changeHolidayCalendarMonth(1)" class="btn btn-sm btn-outline-success">
                        <i class="bi bi-chevron-right"></i>
                    </button>
                </div>
            </div>
            
            <div class="calendar-grid">
        `;
        
        // Day headers
        dayNames.forEach(day => {
            calendarHtml += `<div class="calendar-day-header">${day}</div>`;
        });
        
        // Previous month's days (grayed out)
        for (let i = startingDayOfWeek - 1; i >= 0; i--) {
            const day = daysInPrevMonth - i;
            const date = new Date(year, month - 1, day);
            const dateStr = formatDateForComparison(date);
            const isHolidayDate = isHoliday(dateStr);
            const isWeekend = date.getDay() === 0 || date.getDay() === 6;
            
            calendarHtml += `
                <div class="calendar-day other-month ${isHolidayDate ? 'holiday' : ''} ${isWeekend ? 'weekend' : ''}" 
                     onclick="toggleHoliday('${dateStr}')">
                    ${day}
                </div>
            `;
        }
        
        // Current month's days
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(year, month, day);
            const dateStr = formatDateForComparison(date);
            const isToday = date.toDateString() === today.toDateString();
            const isHolidayDate = isHoliday(dateStr);
            const isWeekend = date.getDay() === 0 || date.getDay() === 6;
            
            let classes = 'calendar-day';
            if (isToday) classes += ' today';
            if (isHolidayDate) classes += ' holiday';
            if (isWeekend) classes += ' weekend';
            
            calendarHtml += `
                <div class="${classes}" onclick="toggleHoliday('${dateStr}')">
                    ${day}
                </div>
            `;
        }
        
        // Next month's days to fill the grid (6 rows total)
        const totalCells = Math.ceil((startingDayOfWeek + daysInMonth) / 7) * 7;
        const remainingCells = totalCells - (startingDayOfWeek + daysInMonth);
        
        for (let day = 1; day <= remainingCells; day++) {
            const date = new Date(year, month + 1, day);
            const dateStr = formatDateForComparison(date);
            const isHolidayDate = isHoliday(dateStr);
            const isWeekend = date.getDay() === 0 || date.getDay() === 6;
            
            calendarHtml += `
                <div class="calendar-day other-month ${isHolidayDate ? 'holiday' : ''} ${isWeekend ? 'weekend' : ''}" 
                     onclick="toggleHoliday('${dateStr}')">
                    ${day}
                </div>
            `;
        }
        
        calendarHtml += `
            </div>
            
            <div class="calendar-legend">
                <div class="legend-item">
                    <div class="legend-color" style="background: #86efac; border: 1px solid #22c55e;"></div>
                    <span>Today</span>
                </div>
                <div class="legend-item">
                    <span style="color: #22c55e; font-size: 12px;">★</span>
                    <span>Holiday</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: #6b7280;"></div>
                    <span>Weekend</span>
                </div>
            </div>
        `;
        
        calendarContainer.innerHTML = calendarHtml;
        
        // Update holiday list
        updateHolidayList();
    }
    
    // Store calendar state globally for navigation
    window.currentHolidayCalendarDate = currentDate;
    window.renderHolidayCalendar = renderHolidayCalendar;
    
    renderHolidayCalendar();
}

// Change holiday calendar month (navigation buttons)
function changeHolidayCalendarMonth(direction) {
    if (!window.currentHolidayCalendarDate) return;
    
    window.currentHolidayCalendarDate.setMonth(window.currentHolidayCalendarDate.getMonth() + direction);
    window.renderHolidayCalendar();
}

// Change holiday calendar to specific month
function changeHolidayCalendarToMonth(monthIndex) {
    if (!window.currentHolidayCalendarDate) return;
    
    window.currentHolidayCalendarDate.setMonth(parseInt(monthIndex));
    window.renderHolidayCalendar();
}

// Change holiday calendar to specific year
function changeHolidayCalendarToYear(year) {
    if (!window.currentHolidayCalendarDate) return;
    
    window.currentHolidayCalendarDate.setFullYear(parseInt(year));
    window.renderHolidayCalendar();
}

// Update holiday list display
function updateHolidayList() {
    const holidayList = document.getElementById('holidayList');
    if (!holidayList) return;
    
    const currentYear = window.currentHolidayCalendarDate?.getFullYear() || new Date().getFullYear();
    const currentMonth = (window.currentHolidayCalendarDate?.getMonth() || new Date().getMonth()) + 1;
    
    holidayList.innerHTML = generateHolidayList(currentYear, currentMonth);
}

// ===================================
// MODAL CALENDAR FUNCTIONS
// ===================================

// Initialize modal calendar for daily attendance
function initializeModalCalendar(availableDates, calendarId, attendanceDisplayId, modalId) {
    const calendarContainer = document.getElementById(calendarId);
    
    if (!calendarContainer) {
        console.error('❌ Modal calendar container not found!');
        return;
    }
    
    const today = new Date();
    let currentDate = new Date(today.getFullYear(), today.getMonth(), 1);
    
    // Convert available dates to a Set for quick lookup
    const dateSet = new Set(availableDates);
    
    // Calculate attendance levels for each date
    const attendanceLevels = calculateAttendanceLevels(availableDates);

    function renderModalCalendar() {
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();
        
        // Get first day of month and number of days
        const firstDay = new Date(year, month, 1);
        const lastDay = new Date(year, month + 1, 0);
        const daysInMonth = lastDay.getDate();
        const startingDayOfWeek = firstDay.getDay();
        
        // Get previous month's last days to fill the grid
        const prevMonth = new Date(year, month, 0);
        const daysInPrevMonth = prevMonth.getDate();
        
        const monthNames = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ];
        
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        
        let calendarHtml = `
            <div class="modal-calendar-header" style="margin-bottom: 1rem;">
                <div class="calendar-navigation">
                    <button onclick="changeModalCalendarMonth(-1, '${calendarId}', '${attendanceDisplayId}', '${modalId}')" class="btn btn-sm btn-outline-success">
                        <i class="bi bi-chevron-left"></i>
                    </button>
                    <div class="calendar-month-year">
                        <select onchange="changeModalCalendarToMonth(this.value, '${calendarId}', '${attendanceDisplayId}', '${modalId}')" class="month-selector">
                            ${monthNames.map((name, index) => 
                                `<option value="${index}" ${index === month ? 'selected' : ''}>${name}</option>`
                            ).join('')}
                        </select>
                        <select onchange="changeModalCalendarToYear(this.value, '${calendarId}', '${attendanceDisplayId}', '${modalId}')" class="year-selector">
                            ${generateYearOptions(year)}
                        </select>
                    </div>
                    <button onclick="changeModalCalendarMonth(1, '${calendarId}', '${attendanceDisplayId}', '${modalId}')" class="btn btn-sm btn-outline-success">
                        <i class="bi bi-chevron-right"></i>
                    </button>
                </div>
            </div>
            
            <div class="calendar-grid">
        `;
        
        // Day headers
        dayNames.forEach(day => {
            calendarHtml += `<div class="calendar-day-header">${day}</div>`;
        });
        
        // Previous month's days (grayed out)
        for (let i = startingDayOfWeek - 1; i >= 0; i--) {
            const day = daysInPrevMonth - i;
            const date = new Date(year, month - 1, day);
            const dateStr = formatDateForComparison(date);
            const hasData = dateSet.has(dateStr);
            const attendanceLevel = attendanceLevels[dateStr] || 'none';
            const isHolidayDate = isHoliday(dateStr);
            const isFutureDate = date > today;
            const isPassed = date < today && date.toDateString() !== today.toDateString();
            const isWeekend = date.getDay() === 0 || date.getDay() === 6; // Sunday = 0, Saturday = 6
            
            let classes = 'calendar-day other-month';
            if (hasData) classes += ' has-data';
            if (isHolidayDate) classes += ' holiday';
            if (isFutureDate) classes += ' future-date';
            if (isPassed) classes += ' passed';
            if (isWeekend) classes += ' weekend';
            classes += ` attendance-${attendanceLevel}`;
            
            calendarHtml += `
                <div class="${classes}" 
                     onclick="selectModalCalendarDate('${dateStr}', '${attendanceDisplayId}', '${modalId}')">
                    ${day}
                </div>
            `;
        }
        
        // Current month's days
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(year, month, day);
            const dateStr = formatDateForComparison(date);
            const hasData = dateSet.has(dateStr);
            const isToday = date.toDateString() === today.toDateString();
            const attendanceLevel = attendanceLevels[dateStr] || 'none';
            const isHolidayDate = isHoliday(dateStr);
            const isFutureDate = date > today;
            const isPassed = date < today && !isToday;
            const isWeekend = date.getDay() === 0 || date.getDay() === 6; // Sunday = 0, Saturday = 6
            
            let classes = 'calendar-day';
            if (isToday) classes += ' today';
            if (hasData) classes += ' has-data';
            if (isHolidayDate) classes += ' holiday';
            if (isFutureDate) classes += ' future-date';
            if (isPassed) classes += ' passed';
            if (isWeekend) classes += ' weekend';
            classes += ` attendance-${attendanceLevel}`;
            
            calendarHtml += `
                <div class="${classes}" onclick="selectModalCalendarDate('${dateStr}', '${attendanceDisplayId}', '${modalId}')">
                    ${day}
                </div>
            `;
        }
        
        // Next month's days to fill the grid (6 rows total)
        const totalCells = Math.ceil((startingDayOfWeek + daysInMonth) / 7) * 7;
        const remainingCells = totalCells - (startingDayOfWeek + daysInMonth);
        
        for (let day = 1; day <= remainingCells; day++) {
            const date = new Date(year, month + 1, day);
            const dateStr = formatDateForComparison(date);
            const hasData = dateSet.has(dateStr);
            const attendanceLevel = attendanceLevels[dateStr] || 'none';
            const isHolidayDate = isHoliday(dateStr);
            const isFutureDate = date > today;
            const isPassed = date < today && date.toDateString() !== today.toDateString();
            const isWeekend = date.getDay() === 0 || date.getDay() === 6; // Sunday = 0, Saturday = 6
            
            let classes = 'calendar-day other-month';
            if (hasData) classes += ' has-data';
            if (isHolidayDate) classes += ' holiday';
            if (isFutureDate) classes += ' future-date';
            if (isPassed) classes += ' passed';
            if (isWeekend) classes += ' weekend';
            classes += ` attendance-${attendanceLevel}`;
            
            calendarHtml += `
                <div class="${classes}" 
                     onclick="selectModalCalendarDate('${dateStr}', '${attendanceDisplayId}', '${modalId}')">
                    ${day}
                </div>
            `;
        }
        
        calendarHtml += `
            </div>
        `;
        
        calendarContainer.innerHTML = calendarHtml;
    }
    
    // Store modal calendar state globally for navigation
    window.currentModalCalendarDate = currentDate;
    window.modalCalendarAvailableDates = dateSet;
    window.modalCalendarAttendanceLevels = attendanceLevels;
    window.renderModalCalendar = renderModalCalendar;
    
    renderModalCalendar();
}

// Select modal calendar date
function selectModalCalendarDate(dateStr, attendanceDisplayId, modalId) {
    console.log(`📅 Modal calendar date selected: ${dateStr}`);
    
    const selectedDate = new Date(dateStr + 'T00:00:00');
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    selectedDate.setHours(0, 0, 0, 0);
    
    const isWeekend = selectedDate.getDay() === 0 || selectedDate.getDay() === 6; // Sunday = 0, Saturday = 6
    const isFutureDate = selectedDate > today;
    const isHolidayDate = isHoliday(dateStr);
    
    // Check if clicked date is weekend
    if (isWeekend) {
        showFloatingNotification('Weekend', 'Attendance is not tracked on weekends', 'info');
        return;
    }
    
    // Check if clicked date is a holiday
    if (isHolidayDate) {
        showFloatingNotification('Holiday', 'This date is marked as a holiday. Attendance data is not tracked for holidays.', 'warning');
        return;
    }
    
    // Check if clicked date is in future
    if (isFutureDate) {
        showFloatingNotification('Future Date', 'Attendance data is not available for future dates', 'info');
        return;
    }

    // Clear previous selection
    document.querySelectorAll('.modal .calendar-day.selected').forEach(day => {
        day.classList.remove('selected');
    });
    
    // Mark new selection
    const clickedElement = event.target;
    clickedElement.classList.add('selected');
    
    // Load attendance for the selected date
    loadModalDailyAttendanceForDate(dateStr, attendanceDisplayId, modalId);
}

// Show floating notification
function showFloatingNotification(title, message, type = 'info') {
    // Remove existing notifications
    const existingNotifications = document.querySelectorAll('.floating-notification');
    existingNotifications.forEach(notification => notification.remove());
    
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `floating-notification floating-notification-${type}`;
    
    const iconClass = type === 'warning' ? 'bi-exclamation-triangle-fill' : 
                     type === 'error' ? 'bi-x-circle-fill' : 'bi-info-circle-fill';
    
    notification.innerHTML = `
        <div class="floating-notification-content">
            <div class="floating-notification-header">
                <i class="bi ${iconClass} floating-notification-icon"></i>
                <span class="floating-notification-title">${title}</span>
                <button class="floating-notification-close" onclick="removeNotification(this.parentElement.parentElement.parentElement)">
                    <i class="bi bi-x"></i>
                </button>
            </div>
            <div class="floating-notification-message">${message}</div>
        </div>
    `;
    
    // Add swipe functionality
    let startX = 0;
    let currentX = 0;
    let isDragging = false;
    let startTime = 0;
    
    notification.addEventListener('touchstart', (e) => {
        startX = e.touches[0].clientX;
        currentX = startX;
        isDragging = true;
        startTime = Date.now();
        notification.style.transition = 'none';
    }, { passive: true });
    
    notification.addEventListener('touchmove', (e) => {
        if (!isDragging) return;
        
        currentX = e.touches[0].clientX;
        const diffX = currentX - startX;
        
        // Only allow swiping right (positive direction)
        if (diffX > 0) {
            notification.style.transform = `translateX(${diffX}px)`;
            notification.style.opacity = Math.max(0.3, 1 - (diffX / 200));
        }
    }, { passive: true });
    
    notification.addEventListener('touchend', () => {
        if (!isDragging) return;
        isDragging = false;
        
        const diffX = currentX - startX;
        const timeDiff = Date.now() - startTime;
        const velocity = Math.abs(diffX) / timeDiff;
        
        notification.style.transition = 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)';
        
        // Remove if swiped far enough or fast enough
        if (diffX > 100 || velocity > 0.5) {
            removeNotification(notification);
        } else {
            // Snap back to original position
            notification.style.transform = 'translateX(0)';
            notification.style.opacity = '1';
        }
    }, { passive: true });
    
    // Mouse events for desktop
    notification.addEventListener('mousedown', (e) => {
        startX = e.clientX;
        currentX = startX;
        isDragging = true;
        startTime = Date.now();
        notification.style.transition = 'none';
        e.preventDefault();
    });
    
    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        
        currentX = e.clientX;
        const diffX = currentX - startX;
        
        if (diffX > 0) {
            notification.style.transform = `translateX(${diffX}px)`;
            notification.style.opacity = Math.max(0.3, 1 - (diffX / 200));
        }
    });
    
    document.addEventListener('mouseup', () => {
        if (!isDragging) return;
        isDragging = false;
        
        const diffX = currentX - startX;
        const timeDiff = Date.now() - startTime;
        const velocity = Math.abs(diffX) / timeDiff;
        
        notification.style.transition = 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)';
        
        if (diffX > 100 || velocity > 0.5) {
            removeNotification(notification);
        } else {
            notification.style.transform = 'translateX(0)';
            notification.style.opacity = '1';
        }
    });
    
    // Add to body
    document.body.appendChild(notification);
    
    // Animate in
    setTimeout(() => {
        notification.classList.add('show');
    }, 10);
    
    // Auto remove after 4 seconds
    setTimeout(() => {
        if (notification.parentElement) {
            removeNotification(notification);
        }
    }, 4000);
}

// Remove notification with animation
function removeNotification(notification) {
    if (!notification || !notification.parentElement) return;
    
    notification.classList.add('hide');
    setTimeout(() => {
        if (notification.parentElement) {
            notification.remove();
        }
    }, 300);
}

// Load today's attendance in modal
function loadModalTodaysAttendance(attendanceDisplayId, modalId = null) {
    const today = new Date();
    const todayStr = formatDateForComparison(today);
    
    // Check if today has data
    const hasDataForToday = extractedData.some(employee => 
        employee.attendanceData.some(record => 
            record.fullDate && formatDateForComparison(record.fullDate) === todayStr
        )
    );
    
    if (hasDataForToday) {
        console.log('📅 Loading today\'s attendance in modal:', todayStr);
        
        // Clear previous selections first
        document.querySelectorAll('.modal .calendar-day.selected').forEach(day => {
            day.classList.remove('selected');
        });
        
        loadModalDailyAttendanceForDate(todayStr, attendanceDisplayId, modalId);
        
        // Highlight today in calendar
        document.querySelectorAll('.modal .calendar-day.today').forEach(day => {
            day.classList.add('selected');
        });
    } else {
        showModalNoDataMessage(todayStr, attendanceDisplayId, 'today');
    }
}

// Modal calendar navigation functions
function changeModalCalendarMonth(direction, calendarId, attendanceDisplayId, modalId) {
    if (window.currentModalCalendarDate) {
        window.currentModalCalendarDate.setMonth(window.currentModalCalendarDate.getMonth() + direction);
        window.renderModalCalendar();
    }
}

function changeModalCalendarToMonth(monthIndex, calendarId, attendanceDisplayId, modalId) {
    if (window.currentModalCalendarDate) {
        window.currentModalCalendarDate.setMonth(parseInt(monthIndex));
        window.renderModalCalendar();
    }
}

function changeModalCalendarToYear(year, calendarId, attendanceDisplayId, modalId) {
    if (window.currentModalCalendarDate) {
        window.currentModalCalendarDate.setFullYear(parseInt(year));
        window.renderModalCalendar();
    }
}