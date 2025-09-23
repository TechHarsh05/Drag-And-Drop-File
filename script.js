
    // Initialize EmailJS
    emailjs.init("axGZ4QYoXmOwRc-rr");

    // Theme toggle functionality
    const themeToggle = document.getElementById('themeToggle');
    const themeIcon = document.querySelector('.theme-toggle-icon');

    // Check for saved theme preference or default to light mode
    const currentTheme = localStorage.getItem('theme') || 'light';
    document.documentElement.setAttribute('data-theme', currentTheme);
    updateThemeIcon(currentTheme);

    themeToggle.addEventListener('click', () => {
      const currentTheme = document.documentElement.getAttribute('data-theme');
      const newTheme = currentTheme === 'light' ? 'dark' : 'light';

      document.documentElement.setAttribute('data-theme', newTheme);
      localStorage.setItem('theme', newTheme);
      updateThemeIcon(newTheme);
    });

    function updateThemeIcon(theme) {
      if (theme === 'dark') {
        themeIcon.textContent = '🌙';
      } else {
        themeIcon.textContent = '☀️';
      }
    }

    // Global variables for edit functionality
    let isEditMode = false;
    let originalData = [];
    let headerMap = {};
    let originalColumnCount = 0;

    const dropArea = document.getElementById('drop-area');
    ['dragenter', 'dragover'].forEach(event => {
      dropArea.addEventListener(event, e => {
        e.preventDefault();
        e.stopPropagation();
        dropArea.classList.add('hover');
      }, false);
    });

    ['dragleave', 'drop'].forEach(event => {
      dropArea.addEventListener(event, e => {
        e.preventDefault();
        e.stopPropagation();
        dropArea.classList.remove('hover');
      }, false);
    });

    dropArea.addEventListener('drop', handleDrop, false);

    document.getElementById('fileElem').addEventListener('change', (e) => {
      handleFiles(e.target.files);
    });

    function handleDrop(e) {
      const files = e.dataTransfer.files;
      handleFiles(files);
    }

    function handleFiles(files) {
      const file = files[0];
      const reader = new FileReader();

      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Clear any previous messages
        document.getElementById('error-container').innerHTML = '';

        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        // Store original data for editing
        originalData = JSON.parse(JSON.stringify(json));
        originalColumnCount = json[0].length;

        headerMap = json[0]?.reduce((map, header, idx) => {
          if (header) map[header.trim().toLowerCase()] = idx;
          return map;
        }, {}) || {};

        // Hide drop zone, show controls
        document.getElementById('drop-heading').style.display = 'none';
        document.getElementById('drop-area').style.display = 'none';
        document.getElementById('buttonContainer').style.display = 'flex';
        document.getElementById('customFilter').style.display = 'block';

        // Generate table
        generateTable(json);
      };

      reader.readAsArrayBuffer(file);
    }

    function generateTable(json) {
      const output = document.getElementById('output');
      let html = '<table><thead><tr>';

      // Original headers
      json[0].forEach((header) => {
        html += `<th>${header ?? ''}</th>`;
      });

      // Calculated headers
      html += `
        <th>Min SP</th>
        <th>Max SP</th>
        <th>Discount %</th>
        <th>Min Price for Everyday Deal</th>
        <th>Max Price for everyday deal</th>
        <th>Waiver %</th>  <!-- Changed from "Discount - Waiver" to "Waiver %" -->
        <th>Status</th>
        <th>Max Price of Best Deal</th>
        <th>Min Price of Best Deal</th>
      `;

      html += '</tr></thead><tbody>';

      // Column indices
      const calculatedFeeIndex = findColumnIndex(headerMap, ['calculated fee', 'calculated', 'fee']);
      const spIndex = findColumnIndex(headerMap, ['s.p', 'sp', 'selling price', 'selling']);
      const dealPriceIndex = findColumnIndex(headerMap, ['deal price', 'everyday deal', 'deal', 'price', 'd.p', 'bau deal price']);
      const waiverIndex = findColumnIndex(headerMap, ['waiver', 'prime waiver', 'waiver amount']);

      for (let rowIndex = 1; rowIndex < json.length; rowIndex++) {
        const row = json[rowIndex];
        let rowHtml = '';
        let rowClass = [];

        // Read values
        const hasCal = calculatedFeeIndex !== -1;
        const hasSP = spIndex !== -1;
        const hasDP = dealPriceIndex !== -1;
        const hasWav = waiverIndex !== -1;
        const rawDP = hasDP ? row[dealPriceIndex] : '';
        
        // Updated condition to check for N/A, empty, or 0
        const isDPZeroOrBlank = hasDP && (rawDP === '' || rawDP === 'N/A' || Number(rawDP) === 0);
        const calculatedFee = hasCal ? parseFloat(row[calculatedFeeIndex]) : NaN;
        const sp = hasSP ? parseFloat(row[spIndex]) : NaN;
        const dealPrice = hasDP ? parseFloat(row[dealPriceIndex]) : NaN;
        const waiver = hasWav ? (parseFloat(row[waiverIndex]) || 0) : NaN;

        // ------- Build computed cells -------
        // Min/Max SP
        const minSPCell = hasCal && !isNaN(calculatedFee)
          ? (calculatedFee * 1.05).toFixed(2)
          : (!hasCal ? "Not found Calculated Fee" : 'N/A');
        const maxSPCell = hasCal && !isNaN(calculatedFee)
          ? (calculatedFee * 1.15).toFixed(2)
          : (!hasCal ? "Not found Calculated Fee" : 'N/A');

        // Discount %
        let discountPercentCell = 'N/A';
        let discountStyle = '';
        if (!hasSP) {
          discountPercentCell = "Not found S.P";
        } else if (!hasDP) {
          discountPercentCell = "Not found Deal Price";
        } else if (isDPZeroOrBlank) {
          discountPercentCell = 'N/A';
        } else if (!isNaN(sp) && sp > 0 && !isNaN(dealPrice)) {
          const discountValue = ((sp - dealPrice) / sp) * 100;
          discountPercentCell = discountValue.toFixed(2) + '%';
          if (discountValue < 5) {
            discountStyle = ' style="background-color: var(--low-discount-bg);"';
            rowClass.push('low-discount');
          }
        }

        // Everyday deal (Min/Max)
        let minEverydayDeal = 'N/A';
        let maxEverydayDeal = 'N/A';
        let minEverydayStyle = '';
        let maxEverydayStyle = '';
        let floor, target, feasible = false;
        if (!hasSP) {
          minEverydayDeal = "Not found S.P";
          maxEverydayDeal = "Not found S.P";
        } else if (!hasCal) {
          minEverydayDeal = "Not found Calculated Fee";
          maxEverydayDeal = "Not found Calculated Fee";
        } else if (!isNaN(sp) && sp > 0 && !isNaN(calculatedFee)) {
          floor = calculatedFee * 0.96;
          target = sp * 0.95;
          feasible = target >= floor;
          if (feasible) {
            minEverydayDeal = floor.toFixed(2);
            maxEverydayDeal = target.toFixed(2);
          } else {
            const msg = 'Cannot provide everyday deal';
            minEverydayDeal = msg;
            maxEverydayDeal = msg;
            minEverydayStyle = ' class="cannot-provide"';
            maxEverydayStyle = ' class="cannot-provide"';
            rowClass.push('low-waiver');
          }
        }

        // Waiver % (changed from Discount - Waiver)
        let waiverPercentCell = 'N/A';
        let waiverPercentValue = NaN; // Store the actual percentage value for calculation
        if (!hasWav) {
          waiverPercentCell = "Not found Waiver";
        } else if (!hasSP) {
          waiverPercentCell = "Not found S.P";
        } else if (isNaN(waiver) || isNaN(sp) || sp === 0) {
          waiverPercentCell = 'N/A';
        } else {
          // Calculate waiver percentage: (waiver * 100) / sp
          waiverPercentValue = (waiver * 100) / sp;
          waiverPercentCell = waiverPercentValue.toFixed(2) + '%';
        }

        // Status - Updated rule: "Too Low" when deal price is less than calculated fee
        let statusCell = '';
        if (!hasCal) {
          statusCell = "Not found Calculated Fee";
        } else if (!hasDP) {
          statusCell = "Not found Deal Price";
        } else if (isDPZeroOrBlank) {
          statusCell = '';
        } else if (!isNaN(calculatedFee) && !isNaN(dealPrice)) {
          // Updated rule: Check if deal price is less than calculated fee
          if (dealPrice < calculatedFee) {
            rowClass.push('low-waiver');
            statusCell = `<span style="background-color:var(--cannot-provide-bg); display:inline-block; padding:2px 6px; border-radius:4px;">Too Low</span>`;
          }
        }

        // Best Deal (Min/Max) - Updated logic
        let minPriceOfBestDeal = 'N/A';
        let maxPriceOfBestDeal = 'N/A';
        let minPriceStyle = '';
        let maxPriceStyle = '';
        let minPriceNumeric = null; // Store numeric value for comparison

        // Check if deal price is 0 or N/A (updated condition)
        if (isDPZeroOrBlank) {
          minPriceOfBestDeal = "Cannot provide best deal";
          maxPriceOfBestDeal = "Cannot provide best deal";
          minPriceStyle = ' class="cannot-provide"';
          maxPriceStyle = ' class="cannot-provide"';
        } else {
          // Calculate Min Price of Best Deal
          if (!hasCal) {
            minPriceOfBestDeal = "Not found Calculated Fee";
          } else if (!hasSP) {
            minPriceOfBestDeal = "Not found S.P";
          } else if (!hasWav) {
            minPriceOfBestDeal = "Not found Waiver";
          } else if (!isNaN(calculatedFee) && !isNaN(waiverPercentValue)) {
            // Calculate minPrice: ((calculated fee - (waiver %)) - 4%)
            const waiverAmount = calculatedFee * (waiverPercentValue / 100);
            const afterWaiver = calculatedFee - waiverAmount;
            const fourPercent = afterWaiver * 0.04;
            const minPrice = afterWaiver - fourPercent;

            if (minPrice > 0) {
              minPriceOfBestDeal = minPrice.toFixed(2);
              minPriceNumeric = minPrice; // Store numeric value
            } else {
              minPriceOfBestDeal = 'Cannot provide best deal';
              minPriceStyle = ' class="cannot-provide"';
            }
          } else {
            minPriceOfBestDeal = 'N/A';
          }

          // Calculate Max Price of Best Deal - Updated with new rule
          if (!hasDP) {
            maxPriceOfBestDeal = "Not found Deal Price";
          } else if (!hasCal) {
            maxPriceOfBestDeal = "Not found Calculated Fee";
          } else if (!isNaN(dealPrice) && !isNaN(calculatedFee)) {
            // New rule: if deal price is less than calculated fee, then treat SP as deal price
            let effectiveDealPrice = dealPrice;
            if (dealPrice < calculatedFee) {
              if (!isNaN(sp)) {
                effectiveDealPrice = sp;
              } else {
                maxPriceOfBestDeal = "Cannot provide best deal (SP not available)";
                maxPriceStyle = ' class="cannot-provide"';
              }
            }

            // Only proceed if we haven't set an error above
            if (maxPriceOfBestDeal !== "Cannot provide best deal (SP not available)") {
              const maxPrice = effectiveDealPrice * 0.95;
              
              // Compare with min price if available
              if (minPriceNumeric !== null && maxPrice < minPriceNumeric) {
                maxPriceOfBestDeal = "Cannot provide best deal";
                maxPriceStyle = ' class="cannot-provide"';
              } else {
                maxPriceOfBestDeal = maxPrice.toFixed(2);
              }
            }
          } else {
            maxPriceOfBestDeal = 'N/A';
          }
        }

        // Row flags
        if (hasSP && hasCal && !isNaN(sp) && !isNaN(calculatedFee)) {
          if (sp < calculatedFee * 1.05) rowClass.push('low-sp');
          else if (sp > calculatedFee * 1.15) rowClass.push('high-sp');
        }

        if (hasDP && !isDPZeroOrBlank && hasSP && hasCal && !isNaN(dealPrice) && !isNaN(sp) && !isNaN(calculatedFee)) {
          const floorCheck = calculatedFee * 0.96;
          const targetCheck = sp * 0.95;
          if (dealPrice < floorCheck || dealPrice > targetCheck) {
            rowClass.push('error-dp');
          }
        }

        if (hasDP && isDPZeroOrBlank) {
          rowClass.push('na-dp');
        }

        if (rowClass.length === 0) rowClass.push('no-error');

        // Render Row
        rowHtml += `<tr class="${rowClass.join(' ')}" data-row="${rowIndex}">`;

        // Original cells
        row.forEach((cell, colIdx) => {
          let displayValue = (cell === '' || cell == null || cell === NaN) ? 'N/A' : cell;
          let style = '';

          if (colIdx === dealPriceIndex && hasDP && isDPZeroOrBlank) {
            displayValue = 'N/A';
          }

          if (displayValue !== 'N/A') {
            if (colIdx === spIndex && hasSP && hasCal && !isNaN(sp) && !isNaN(calculatedFee)) {
              const minSPValue = calculatedFee * 1.05;
              const maxSPValue = calculatedFee * 1.15;
              if (sp < minSPValue) style = 'background-color: var(--low-sp-bg);';
              else if (sp > maxSPValue) style = 'background-color: var(--high-sp-bg);';
            }

            if (colIdx === dealPriceIndex && hasDP && !isDPZeroOrBlank && hasSP && hasCal && !isNaN(dealPrice) && !isNaN(sp) && !isNaN(calculatedFee)) {
              const floorCheck = calculatedFee * 0.96;
              const targetCheck = sp * 0.95;
              if (dealPrice < floorCheck || dealPrice > targetCheck) {
                style = 'background-color: var(--error-dp-bg);';
              }
            }
          }

          rowHtml += `<td style="${style}" data-col="${colIdx}">${displayValue}</td>`;
        });

        // Calculated cells
        rowHtml += `
          <td>${minSPCell}</td>
          <td>${maxSPCell}</td>
          <td${discountStyle}>${discountPercentCell}</td>
          <td${minEverydayStyle}>${minEverydayDeal}</td>
          <td${maxEverydayStyle}>${maxEverydayDeal}</td>
          <td>${waiverPercentCell}</td>
          <td>${statusCell || ''}</td>
          <td${maxPriceStyle}>${maxPriceOfBestDeal}</td>
          <td${minPriceStyle}>${minPriceOfBestDeal}</td>
        `;

        rowHtml += '</tr>';
        html += rowHtml;
      }

      html += '</tbody></table>';
      output.innerHTML = html;

      // Add edit button if not already present
      if (!document.getElementById('editBtn')) {
        const editBtn = document.createElement('button');
        editBtn.id = 'editBtn';
        editBtn.className = 'action-button';
        editBtn.textContent = 'Edit';
        editBtn.onclick = toggleEditMode;

        // Insert edit button before the download button
        const downloadBtn = document.getElementById('download');
        downloadBtn.parentNode.insertBefore(editBtn, downloadBtn);
      }
    }

    function toggleEditMode() {
      isEditMode = !isEditMode;
      const editBtn = document.getElementById('editBtn');
      const table = document.querySelector('#output table');

      if (isEditMode) {
        editBtn.textContent = 'Cancel';
        editBtn.classList.add('danger');

        // Make cells editable
        const cells = table.querySelectorAll('tbody td[data-col]');
        cells.forEach(cell => {
          const value = cell.textContent;
          cell.innerHTML = `<input type="text" value="${value}" style="width: 100%; border: none; background: transparent;">`;
        });

        // Add save button
        const saveBtn = document.createElement('button');
        saveBtn.id = 'saveBtn';
        saveBtn.className = 'action-button success';
        saveBtn.textContent = 'Save';
        saveBtn.onclick = saveChanges;

        // Insert save button after edit button
        editBtn.parentNode.insertBefore(saveBtn, editBtn.nextSibling);
      } else {
        editBtn.textContent = 'Edit';
        editBtn.classList.remove('danger');

        // Remove save button if exists
        const saveBtn = document.getElementById('saveBtn');
        if (saveBtn) saveBtn.remove();

        // Regenerate table to discard changes
        generateTable(originalData);
      }
    }

    function saveChanges() {
      const table = document.querySelector('#output table');
      const rows = table.querySelectorAll('tbody tr');

      // Update original data
      rows.forEach(row => {
        const rowIndex = parseInt(row.getAttribute('data-row'));
        const cells = row.querySelectorAll('td[data-col]');

        cells.forEach(cell => {
          const colIndex = parseInt(cell.getAttribute('data-col'));
          const input = cell.querySelector('input');
          if (input) {
            originalData[rowIndex][colIndex] = input.value;
          }
        });
      });

      // Exit edit mode
      isEditMode = false;
      const editBtn = document.getElementById('editBtn');
      editBtn.textContent = 'Edit';
      editBtn.classList.remove('danger');
      document.getElementById('saveBtn').remove();

      // Regenerate table with updated data
      generateTable(originalData);

      // Show success message
      const successMsg = document.createElement('div');
      successMsg.className = 'success-message';
      successMsg.textContent = 'Changes saved successfully!';
      document.getElementById('error-container').appendChild(successMsg);

      // Remove message after 3 seconds
      setTimeout(() => {
        successMsg.remove();
      }, 3000);
    }

    function findColumnIndex(headerMap, possibleNames) {
      for (const name of possibleNames) {
        if (headerMap[name.toLowerCase()] !== undefined) {
          return headerMap[name.toLowerCase()];
        }
      }
      return -1;
    }

    function applyCustomFilter() {
      const filter = document.getElementById("customFilter").value;
      const rows = document.querySelectorAll("#output table tbody tr");
      rows.forEach((row) => {
        row.style.display = (filter === "all" || row.classList.contains(filter)) ? "" : "none";
      });
    }

    function downloadExcel() {
      const table = document.querySelector('#output table');
      if (!table) return alert("No table to export!");

      const clonedTable = table.cloneNode(true);
      const rows = clonedTable.querySelectorAll("tr");
      rows.forEach((row, index) => {
        if (index !== 0 && row.style.display === "none") row.remove();
      });

      const worksheet = XLSX.utils.table_to_sheet(clonedTable);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered Data");
      XLSX.writeFile(workbook, "Filtered_Data.xlsx");
    }

    function openModal() {
      document.getElementById('contactModal').style.display = 'block';
    }

    function closeModal() {
      document.getElementById('contactModal').style.display = 'none';
    }

    document.getElementById('contactForm').addEventListener('submit', function (e) {
      e.preventDefault();
      emailjs.sendForm('service_ivdpwoe', 'template_flcizdb', this)
        .then(() => {
          alert('Message sent successfully!');
          closeModal();
          this.reset();
        }, (err) => {
          alert('Failed to send message. Try again.');
          console.error(err);
        });
    });