// Main variables
let attendanceData = [];
let subjects = [];
let timerInterval;
let timerRunning = false;
let remainingTime = 25 * 60; // Default 25 minutes
let scene, camera, renderer, cube;

// DOM Elements
document.addEventListener('DOMContentLoaded', function() {
    // Initialize the app
    initApp();
});

function initApp() {
    // Set current date
    updateCurrentDate();
    
    // Load data from local storage
    loadData();
    
    // Initialize navigation
    initNavigation();
    
    // Initialize attendance table
    updateAttendanceTable();
    
    // Initialize analytics
    updateAnalytics();
    
    // Initialize 3D visual
    init3DVisual();
    
    // Initialize event listeners
    initEventListeners();
}

// Data Management Functions
function loadData() {
    // Load subjects
    const savedSubjects = localStorage.getItem('subjects');
    if (savedSubjects) {
        subjects = JSON.parse(savedSubjects);
    } else {
        // Default subjects if none exist
        subjects = ['Mathematics', 'Physics', 'Chemistry', 'English'];
        saveSubjects();
    }
    
    // Load attendance data
    const savedData = localStorage.getItem('attendanceData');
    if (savedData) {
        attendanceData = JSON.parse(savedData);
    }
    
    // Check if we need to add today's entry
    const today = new Date().toISOString().split('T')[0];
    const todayEntry = attendanceData.find(entry => entry.date === today);
    
    if (!todayEntry) {
        addNewDateEntry(today);
    }
}

function saveData() {
    localStorage.setItem('attendanceData', JSON.stringify(attendanceData));
    showNotification('Data saved successfully!', 'success');
}

function saveSubjects() {
    localStorage.setItem('subjects', JSON.stringify(subjects));
}

function addNewDateEntry(date) {
    const dayOfWeek = new Date(date).getDay();
    const isHoliday = dayOfWeek === 0; // Sunday
    
    const newEntry = {
        date: date,
        isHoliday: isHoliday
    };
    
    // Add attendance status for each subject
    subjects.forEach(subject => {
        newEntry[subject] = isHoliday ? 'holiday' : null;
    });
    
    attendanceData.push(newEntry);
    saveData();
}

function exportToExcel() {
    try {
        // Format data for Excel
        const wsData = [
            ['Date', ...subjects]
        ];
        
        attendanceData.forEach(entry => {
            const row = [entry.date];
            subjects.forEach(subject => {
                if (entry.isHoliday) {
                    row.push('Holiday');
                } else {
                    row.push(entry[subject] === 'present' ? 'Present' : 
                           entry[subject] === 'absent' ? 'Absent' : 'N/A');
                }
            });
            wsData.push(row);
        });
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        // Create workbook
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Attendance');
        
        // Generate Excel file
        XLSX.writeFile(wb, 'attendance_report.xlsx');
        
        showNotification('Excel file downloaded successfully!', 'success');
    } catch (error) {
        showNotification('Error exporting to Excel: ' + error.message, 'error');
    }
}

function clearAllData() {
    if (confirm("Are you sure you want to clear all attendance data? This cannot be undone.")) {
        attendanceData = [];
        // Add today's entry after clearing
        const today = new Date().toISOString().split('T')[0];
        addNewDateEntry(today);
        
        saveData();
        updateAttendanceTable();
        updateAnalytics();
        showNotification('All data has been cleared', 'success');
    }
}

function downloadBackup() {
    const backupData = {
        subjects: subjects,
        attendanceData: attendanceData,
        timestamp: new Date().toISOString()
    };
    
    const dataStr = JSON.stringify(backupData);
    const dataBlob = new Blob([dataStr], {type: 'application/json'});
    const url = URL.createObjectURL(dataBlob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'attendance_backup_' + new Date().toISOString().split('T')[0] + '.json';
    document.body.appendChild(a);
    a.click();
    
    setTimeout(() => {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }, 0);
    
    showNotification('Backup downloaded successfully!', 'success');
}

function restoreFromBackup(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const backupData = JSON.parse(e.target.result);
            
            // Validate backup data
            if (!backupData.subjects || !backupData.attendanceData) {
                throw new Error('Invalid backup file');
            }
            
            subjects = backupData.subjects;
            attendanceData = backupData.attendanceData;
            
            saveSubjects();
            saveData();
            
            // Refresh UI
            updateAttendanceTable();
            updateSubjectsList();
            updateAnalytics();
            
            showNotification('Backup restored successfully!', 'success');
        } catch (error) {
            showNotification('Error restoring backup: ' + error.message, 'error');
        }
    };
    reader.readAsText(file);
}

// UI Update Functions
function updateCurrentDate() {
    const currentDateElement = document.getElementById('current-date');
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    currentDateElement.textContent = new Date().toLocaleDateString(undefined, options);
}

function updateAttendanceTable() {
    // Update table header with subjects
    const tableHeader = document.getElementById('table-header');
    tableHeader.innerHTML = '<th>Date</th>';
    
    subjects.forEach(subject => {
        tableHeader.innerHTML += `<th>${subject}</th>`;
    });
    
    // Update table body with attendance data
    const tableBody = document.getElementById('attendance-data');
    tableBody.innerHTML = '';
    
    // Sort data by date, most recent first
    const sortedData = [...attendanceData].sort((a, b) => new Date(b.date) - new Date(a.date));
    
    sortedData.forEach(entry => {
        const row = document.createElement('tr');
        
        // Format date
        const dateObj = new Date(entry.date);
        const formattedDate = dateObj.toLocaleDateString(undefined, { 
            year: 'numeric', 
            month: 'short', 
            day: 'numeric',
            weekday: 'short'
        });
        
        row.innerHTML = `<td>${formattedDate}</td>`;
        
        subjects.forEach(subject => {
            if (entry.isHoliday) {
                row.innerHTML += `<td class="status-holiday">Holiday</td>`;
            } else {
                const status = entry[subject];
                if (status === 'present') {
                    row.innerHTML += `<td class="status-present">Present</td>`;
                } else if (status === 'absent') {
                    row.innerHTML += `<td class="status-absent">Absent</td>`;
                } else {
                    row.innerHTML += `
                        <td>
                            <select class="attendance-select" data-date="${entry.date}" data-subject="${subject}">
                                <option value="">Select</option>
                                <option value="present">Present</option>
                                <option value="absent">Absent</option>
                            </select>
                        </td>
                    `;
                }
            }
        });
        
        tableBody.appendChild(row);
    });
    
    // Add event listeners to select elements
    document.querySelectorAll('.attendance-select').forEach(select => {
        select.addEventListener('change', function() {
            const date = this.dataset.date;
            const subject = this.dataset.subject;
            const status = this.value;
            
            // Update attendance data
            const entry = attendanceData.find(e => e.date === date);
            if (entry) {
                entry[subject] = status;
                saveData();
                updateAttendanceTable();
                updateAnalytics();
            }
        });
    });
    
    // Also update the subject dropdown in analytics
    const subjectSelect = document.getElementById('subject-select');
    subjectSelect.innerHTML = '';
    
    subjects.forEach(subject => {
        subjectSelect.innerHTML += `<option value="${subject}">${subject}</option>`;
    });
    
    // Update subject list in settings
    updateSubjectsList();
}

function updateSubjectsList() {
    const subjectsList = document.getElementById('subjects-list');
    subjectsList.innerHTML = '';
    
    subjects.forEach(subject => {
        const subjectItem = document.createElement('div');
        subjectItem.className = 'subject-item';
        subjectItem.innerHTML = `
            <span>${subject}</span>
            <button class="delete-subject" data-subject="${subject}">&times;</button>
        `;
        subjectsList.appendChild(subjectItem);
    });
    
    // Add event listeners to delete buttons
    document.querySelectorAll('.delete-subject').forEach(button => {
        button.addEventListener('click', function() {
            const subject = this.dataset.subject;
            if (subjects.length > 1) {
                deleteSubject(subject);
            } else {
                showNotification('Cannot delete the last subject', 'error');
            }
        });
    });
}

function updateAnalytics() {
    // Update overall attendance chart
    updateOverallChart();
    
    // Update subject chart for the first subject (default)
    if (subjects.length > 0) {
        updateSubjectChart(subjects[0]);
    }
    
    // Add event listener to subject select
    document.getElementById('subject-select').addEventListener('change', function() {
        updateSubjectChart(this.value);
    });
}

function updateOverallChart() {
    const ctx = document.getElementById('overall-chart').getContext('2d');
    
    // Calculate totals
    let totalPresent = 0;
    let totalAbsent = 0;
    let totalPending = 0;
    
    attendanceData.forEach(entry => {
        if (!entry.isHoliday) {
            subjects.forEach(subject => {
                if (entry[subject] === 'present') {
                    totalPresent++;
                } else if (entry[subject] === 'absent') {
                    totalAbsent++;
                } else if (entry[subject] === null) {
                    totalPending++;
                }
            });
        }
    });
    
    // If there's an existing chart, destroy it
    if (window.overallChart) {
        window.overallChart.destroy();
    }
    
    // Create new chart
    window.overallChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['Present', 'Absent', 'Pending'],
            datasets: [{
                data: [totalPresent, totalAbsent, totalPending],
                backgroundColor: [
                    '#28a745',
                    '#d93654',
                    '#6c757d'
                ],
                borderWidth: 2,
                borderColor: '#1e1e2d'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#e6e6e6',
                        font: {
                            size: 12
                        }
                    }
                }
            }
        }
    });
}

function updateSubjectChart(subject) {
    const ctx = document.getElementById('subject-chart').getContext('2d');
    
    // Calculate subject stats
    let present = 0;
    let absent = 0;
    
    attendanceData.forEach(entry => {
        if (!entry.isHoliday) {
            if (entry[subject] === 'present') {
                present++;
            } else if (entry[subject] === 'absent') {
                absent++;
            }
        }
    });
    
    // Update stats
    document.getElementById('present-count').textContent = present;
    document.getElementById('absent-count').textContent = absent;
    
    const totalClasses = present + absent;
    const percentage = totalClasses > 0 ? Math.round((present / totalClasses) * 100) : 0;
    document.getElementById('attendance-percentage').textContent = percentage + '%';
    
    // Calculate classes needed to reach 75% attendance
    let classesToAttend = 0;
    if (percentage < 75) {
        // Formula: (0.75 * (present + absent + x) - present) = x
        classesToAttend = Math.ceil((0.75 * (totalClasses) - present) / 0.25);
    }
    document.getElementById('classes-to-attend').textContent = classesToAttend;
    
    // If there's an existing chart, destroy it
    if (window.subjectChart) {
        window.subjectChart.destroy();
    }
    
    // Create new chart
    window.subjectChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Present', 'Absent'],
            datasets: [{
                data: [present, absent],
                backgroundColor: [
                    '#28a745',
                    '#d93654'
                ],
                borderWidth: 2,
                borderColor: '#1e1e2d'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#e6e6e6',
                        font: {
                            size: 12
                        }
                    }
                }
            }
        }
    });
}

// Subject Management
function addSubject(subjectName) {
    if (!subjectName.trim()) {
        showNotification('Subject name cannot be empty', 'error');
        return;
    }
    
    if (subjects.includes(subjectName)) {
        showNotification('Subject already exists', 'error');
        return;
    }
    
    subjects.push(subjectName);
    saveSubjects();
    
    // Add the new subject to all attendance entries
    attendanceData.forEach(entry => {
        entry[subjectName] = entry.isHoliday ? 'holiday' : null;
    });
    
    saveData();
    updateAttendanceTable();
    updateAnalytics();
    
    showNotification(`Subject "${subjectName}" added successfully`, 'success');
}

function deleteSubject(subjectName) {
    const index = subjects.indexOf(subjectName);
    if (index !== -1) {
        subjects.splice(index, 1);
        saveSubjects();
        
        // Remove the subject from all attendance entries
        attendanceData.forEach(entry => {
            delete entry[subjectName];
        });
        
        saveData();
        updateAttendanceTable();
        updateAnalytics();
        
        showNotification(`Subject "${subjectName}" deleted`, 'success');
    }
}

// Timer Functions
function updateTimerDisplay() {
    const minutes = Math.floor(remainingTime / 60);
    const seconds = remainingTime % 60;
    
    document.getElementById('minutes').textContent = minutes.toString().padStart(2, '0');
    document.getElementById('seconds').textContent = seconds.toString().padStart(2, '0');
}

function startTimer() {
    if (!timerRunning) {
        timerRunning = true;
        document.getElementById('start-timer').disabled = true;
        document.getElementById('pause-timer').disabled = false;
        
        timerInterval = setInterval(() => {
            remainingTime--;
            
            // Animate the cube faster as time runs out
            rotateCube(remainingTime);
            
            if (remainingTime <= 0) {
                clearInterval(timerInterval);
                timerRunning = false;
                document.getElementById('start-timer').disabled = false;
                document.getElementById('pause-timer').disabled = true;
                
                // Play alarm sound
                playAlarmSound();
                
                showNotification('Timer finished!', 'success');
            }
            
            updateTimerDisplay();
        }, 1000);
    }
}

function pauseTimer() {
    if (timerRunning) {
        clearInterval(timerInterval);
        timerRunning = false;
        document.getElementById('start-timer').disabled = false;
        document.getElementById('pause-timer').disabled = true;
    }
}

function resetTimer() {
    clearInterval(timerInterval);
    timerRunning = false;
    remainingTime = 25 * 60; // Reset to 25 minutes
    updateTimerDisplay();
    document.getElementById('start-timer').disabled = false;
    document.getElementById('pause-timer').disabled = true;
}

function setTimer(minutes) {
    clearInterval(timerInterval);
    timerRunning = false;
    remainingTime = minutes * 60;
    updateTimerDisplay();
    document.getElementById('start-timer').disabled = false;
    document.getElementById('pause-timer').disabled = true;
}

function playAlarmSound() {
    // Create audio element
    const audio = new Audio('https://soundbible.com/grab.php?id=1252&type=mp3');
    audio.play().catch(e => {
        // Fallback for browsers that block autoplay
        showNotification('Timer finished! (Sound was blocked by browser)', 'success');
    });
}

// 3D Visual Functions
function init3DVisual() {
    // Create scene
    scene = new THREE.Scene();
    scene.background = new THREE.Color(0x1a1a2e);
    
    // Create camera
    camera = new THREE.PerspectiveCamera(75, 1, 0.1, 1000);
    camera.position.z = 5;
    
    // Create renderer
    const container = document.getElementById('visual-container');
    renderer = new THREE.WebGLRenderer({ antialias: true });
    renderer.setSize(container.clientWidth, container.clientHeight);
    container.appendChild(renderer.domElement);
    
    // Create a cube
    const geometry = new THREE.BoxGeometry(2, 2, 2);
    const material = new THREE.MeshBasicMaterial({ 
        color: 0x6d5acd,
        wireframe: true
    });
    
    cube = new THREE.Mesh(geometry, material);
    scene.add(cube);
    
    // Start animation
    animate();
    
    // Handle window resize
    window.addEventListener('resize', () => {
        if (container.clientWidth > 0 && container.clientHeight > 0) {
            renderer.setSize(container.clientWidth, container.clientHeight);
            camera.aspect = container.clientWidth / container.clientHeight;
            camera.updateProjectionMatrix();
        }
    });
}

function animate() {
    requestAnimationFrame(animate);
    
    cube.rotation.x += 0.005;
    cube.rotation.y += 0.01;
    
    renderer.render(scene, camera);
}

function rotateCube(time) {
    // Make rotation speed increase as timer runs down
    const intensity = Math.max(0.5, (25 * 60 - time) / (25 * 60)) * 2;
    cube.rotation.x += 0.005 * intensity;
    cube.rotation.y += 0.01 * intensity;
}

// Helper Functions
function showNotification(message, type) {
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    // Animate in
    setTimeout(() => {
        notification.style.right = '20px';
    }, 10);
    
    // Remove after 3 seconds
    setTimeout(() => {
        notification.style.right = '-300px';
        setTimeout(() => {
            document.body.removeChild(notification);
        }, 300);
    }, 3000);
}

// Navigation Functions
function initNavigation() {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', function() {
            const sectionId = this.dataset.section;
            
            // Update active nav item
            document.querySelectorAll('.nav-item').forEach(navItem => {
                navItem.classList.remove('active');
            });
            this.classList.add('active');
            
            // Show corresponding section
            document.querySelectorAll('.content-section').forEach(section => {
                section.classList.remove('active');
            });
            document.getElementById(sectionId).classList.add('active');
            
            // Special handling for sections
            if (sectionId === 'analysis') {
                updateAnalytics();
            } else if (sectionId === 'timer') {
                // Reset renderer size when switching to timer section
                setTimeout(() => {
                    const container = document.getElementById('visual-container');
                    if (renderer && container) {
                        renderer.setSize(container.clientWidth, container.clientHeight);
                        camera.aspect = container.clientWidth / container.clientHeight;
                        camera.updateProjectionMatrix();
                    }
                }, 100);
            }
        });
    });
}

// Initialize all event listeners
function initEventListeners() {
    // Save data button
    document.getElementById('save-data-btn').addEventListener('click', saveData);
    
    // Export to Excel button
    document.getElementById('export-excel-btn').addEventListener('click', exportToExcel);
    
    // Add subject button (table)
    document.getElementById('add-subject-btn').addEventListener('click', () => {
        document.getElementById('add-subject-modal').classList.add('show');
    });
    
    // Add subject button (settings)
    document.getElementById('add-subject-settings').addEventListener('click', () => {
        const subjectName = document.getElementById('new-subject-input').value;
        if (subjectName.trim()) {
            addSubject(subjectName);
            document.getElementById('new-subject-input').value = '';
        }
    });
    
    // Modal close button
    document.querySelector('.close-modal').addEventListener('click', () => {
        document.getElementById('add-subject-modal').classList.remove('show');
    });
    
    // Modal cancel button
    document.getElementById('cancel-subject').addEventListener('click', () => {
        document.getElementById('add-subject-modal').classList.remove('show');
    });
    
    // Modal save button
    document.getElementById('save-subject').addEventListener('click', () => {
        const subjectName = document.getElementById('subject-name').value;
        if (subjectName.trim()) {
            addSubject(subjectName);
            document.getElementById('subject-name').value = '';
            document.getElementById('add-subject-modal').classList.remove('show');
        }
    });
    
    // Clear data button
    document.getElementById('clear-data').addEventListener('click', clearAllData);
    
    // Download backup button
    document.getElementById('download-backup').addEventListener('click', downloadBackup);
    
    // Restore file input
    document.getElementById('restore-file').addEventListener('change', function() {
        if (this.files.length > 0) {
            restoreFromBackup(this.files[0]);
        }
    });
    
    // Timer buttons
    document.getElementById('start-timer').addEventListener('click', startTimer);
    document.getElementById('pause-timer').addEventListener('click', pauseTimer);
    document.getElementById('reset-timer').addEventListener('click', resetTimer);
    
    // Timer preset buttons
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const minutes = parseInt(this.dataset.time);
            setTimer(minutes);
        });
    });
    
    // Add CSS for notifications
    const style = document.createElement('style');
    style.textContent = `
        .notification {
            position: fixed;
            right: -300px;
            top: 20px;
            background-color: #333;
            color: white;
            padding: 15px 20px;
            border-radius: 5px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.2);
            z-index: 1000;
            transition: right 0.3s ease;
        }
        
        .notification.success {
            background-color: #28a745;
        }
        
        .notification.error {
            background-color: #d93654;
        }
    `;
    document.head.appendChild(style);
}