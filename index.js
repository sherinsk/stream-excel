const express = require('express');
const { PrismaClient } = require('@prisma/client');
const ExcelJS = require('exceljs');
const cors = require('cors');

const prisma = new PrismaClient();
const app = express();

app.use(cors());

app.get('/download-excel', async (req, res) => {
    try {
      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        stream: res,
      });
  
      const worksheet = workbook.addWorksheet('YourTable Data');
  
      worksheet.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 20 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Phone Number', key: 'phoneNumber', width: 15 },
        { header: 'Age', key: 'age', width: 10 },
        { header: 'Batch', key: 'batch', width: 15 },
      ];
  
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=YourTableData.xlsx');
  
      let cursor = null;
      const chunkSize = 1000; // Increased chunk size for faster streaming
      let chunkCount = 0;
  
      console.log('Starting data streaming...');
  
      while (true) {
        const data = await prisma.yourTable.findMany({
          take: chunkSize,
          skip: cursor ? 1 : 0,
          cursor: cursor ? { id: cursor.id } : undefined,
        });
  
        console.log(`Processing chunk ${++chunkCount}, rows: ${data.length}`);
  
        if (data.length === 0) {
          break; // Exit the loop when no more data is fetched
        }
  
        data.forEach((row) => {
          worksheet.addRow(row).commit();
        });
  
        cursor = data[data.length - 1];
      }
  
      console.log('Finalizing workbook...');
      await workbook.commit();
      console.log('Excel file generation completed.');
    } catch (error) {
      console.error('Error generating Excel file:', error);
      res.status(500).send('An error occurred while generating the Excel file.');
    }
  });  

app.listen(3000, () => {
  console.log('Server is running on port 3000');
});
