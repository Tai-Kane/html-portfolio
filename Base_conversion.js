// 16進制卡號每2碼倒轉
        function reverseHexPairs(hexNum) {
          if (hexNum.length % 2 !== 0) {
              return "無效的16進制卡號，長度應為偶數";
          }
          
          let reversedPairs = [];
          for (let i = 0; i < hexNum.length; i += 2) {
              reversedPairs.push(hexNum.slice(i, i + 2));
          }
          reversedPairs.reverse();  // 將所有對調
          return reversedPairs.join('');  // 重新合併為一個字符串
      }

      // 16進制轉10進制
      function hexToDecimal(hexNum) {
          try {
              let decimalNum = parseInt(hexNum, 16);
              return decimalNum.toString().padStart(hexNum.length, '0');
          } catch (error) {
              return "無效的16進制數字";
          }
      }

      // 主處理函數
      function processHex() {
          let hexNum = document.getElementById("hexNum").value.trim().toUpperCase();

          if (!hexNum) {
              document.getElementById("result").innerHTML = "請輸入16進制卡號";
              return;
          }

          if (hexNum.length < 8) {
              document.getElementById("result").innerHTML = "未達8個字元 請輸入完整";
              return;
          }

          if (hexNum.length > 8) {
              document.getElementById("result").innerHTML = "超過8個字元 請重新確認";
              return;
          }

          // 倒轉卡號
          let reversedNum = reverseHexPairs(hexNum);
          let decimalNum = hexToDecimal(reversedNum);

          // 顯示結果
          document.getElementById("result").innerHTML = `
              <strong>卡號倒轉：</strong> ${reversedNum} <br>
              <strong>10進制卡號：</strong> ${decimalNum}
          `;
      }

      function processExcel() {
          let fileInput = document.getElementById('excelFile');
          let file = fileInput.files[0];
      
          if (!file) {
              document.getElementById("result1").innerHTML = "請上傳EXCEL文件";
              return;
          }
      
          let reader = new FileReader();
          reader.onload = function(e) {
              let data = new Uint8Array(e.target.result);
              let workbook = XLSX.read(data, { type: 'array' });
              let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
              let rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
      
              let results = rows.map(row => {
                  let hexNum = row[0].toString().trim().toUpperCase();
                  if (hexNum.length < 8) {
                      return {
                          original: hexNum,
                          error: "original未達8個字元"
                      };
                  }
                  if (hexNum.length > 8) {
                      return {
                          original: hexNum,
                          error: "original超過8個字元"
                      };
                  }
                  let reversedNum = reverseHexPairs(hexNum);
                  let decimalNum = hexToDecimal(reversedNum);
                  return {
                      original: hexNum,
                      decimal: decimalNum
                  };
              });
      
              exportToExcel(results);
          };
          reader.readAsArrayBuffer(file);
      }
      
      function exportToExcel(results) {
          let worksheet = XLSX.utils.json_to_sheet(results);
          let workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, "Results");
      
          XLSX.writeFile(workbook, "轉換結果.xlsx");
      }
