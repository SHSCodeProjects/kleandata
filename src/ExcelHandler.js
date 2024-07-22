import React, { useState, useRef } from 'react';
import ExcelJS from 'exceljs';
import {
  Button, IconButton, Tabs, Tab, Box, Paper, Table, TableHead, TableRow, TableCell, TableBody, Tooltip, TextField, Dialog, DialogTitle, DialogContent, DialogActions, Select, MenuItem,
} from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import MergeIcon from '@mui/icons-material/Merge';
import AddIcon from '@mui/icons-material/Add';

const ExcelHandler = () => {
  const [workbook, setWorkbook] = useState(null);
  const [data, setData] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState(0);
  const [editCell, setEditCell] = useState({ rowIndex: null, columnKey: null, value: '' });
  const [newRow, setNewRow] = useState({});
  const [mergeDialogOpen, setMergeDialogOpen] = useState(false);
  const [currentRowIndex, setCurrentRowIndex] = useState(null);
  const [targetSheetIndex, setTargetSheetIndex] = useState(null);
  const [manualSelectDialogOpen, setManualSelectDialogOpen] = useState(false);
  const [manualSelectOptions, setManualSelectOptions] = useState([]);
  const [selectedManualRow, setSelectedManualRow] = useState(null);
  const [undoStack, setUndoStack] = useState([]);
  const [isUndoDisabled, setIsUndoDisabled] = useState(true);
  const tableRef = useRef(null);

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = async (e) => {
      const buffer = e.target.result;
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      setWorkbook(wb);
      loadSheetData(wb.worksheets[0]);
    };
    reader.readAsArrayBuffer(file);
  };

  const loadSheetData = (sheet) => {
    const headers = sheet.getRow(1).values.slice(1); // Ignore first element
    const jsonData = sheet.getSheetValues().slice(2).map((row) => {
      const rowData = {};
      headers.forEach((header, i) => {
        rowData[header] = row[i + 1];
      });
      return rowData;
    });
    setData(jsonData);
  };

  const handleSheetChange = (event, newValue) => {
    setSelectedSheet(newValue);
    const sheet = workbook.worksheets[newValue];
    loadSheetData(sheet);
  };

  const deleteRow = (index) => {
    const updatedData = data.filter((_, i) => i !== index);
    setUndoStack([...undoStack, { type: 'delete', data: data[index], index, sheetIndex: selectedSheet }]);
    setData(updatedData);
    const sheet = workbook.worksheets[selectedSheet];
    sheet.spliceRows(index + 2, 1);
    setIsUndoDisabled(false);
  };

  const handleMergeClick = (rowIndex) => {
    setCurrentRowIndex(rowIndex);
    setMergeDialogOpen(true);
  };

  const handleMerge = async () => {
    const selectedRow = data[currentRowIndex];
    if (!selectedRow) return;

    const targetSheet = workbook.worksheets[targetSheetIndex];
    if (!targetSheet) return;

    const targetHeaders = targetSheet.getRow(1).values.slice(1);
    const targetFirstNameIndex = targetHeaders.indexOf('firstName') + 1;
    const targetLastNameIndex = targetHeaders.indexOf('lastName') + 1;

    let matchingRow;
    targetSheet.eachRow((row) => {
      const rowFirstName = row.getCell(targetFirstNameIndex).value?.toString().toLowerCase().trim();
      const rowLastName = row.getCell(targetLastNameIndex).value?.toString().toLowerCase().trim();
      const selectedFullNameParts = selectedRow.fullName.toLowerCase().split(' ');
      const selectedFirstName = selectedFullNameParts[0].trim();
      const selectedLastName = selectedFullNameParts[selectedFullNameParts.length - 1].trim();

      if (rowFirstName === selectedFirstName && rowLastName === selectedLastName) {
        matchingRow = row;
      }
    });

    if (!matchingRow) {
      const options = [];
      targetSheet.eachRow((row) => {
        const firstName = row.getCell(targetFirstNameIndex).value;
        const lastName = row.getCell(targetLastNameIndex).value;
        options.push({ row, firstName, lastName });
      });
      setManualSelectOptions(options);
      setManualSelectDialogOpen(true);
      return;
    }

    await mergeRows(matchingRow, selectedRow, targetHeaders, targetSheet);
    setMergeDialogOpen(false);
  };

  const mergeRows = async (matchingRow, selectedRow, targetHeaders, targetSheet) => {
    const firstEmptyCol = matchingRow.cellCount + 1;
    const originalRowValues = { ...matchingRow.values };
    Object.keys(selectedRow).forEach((key) => {
      if (key !== 'firstName' && key !== 'lastName' && key !== 'fullName') {
        const columnIndex = targetHeaders.indexOf(key) + 1;
        if (columnIndex > 0) {
          matchingRow.getCell(columnIndex).value = selectedRow[key];
        } else {
          const newColumnIndex = firstEmptyCol + targetHeaders.length;
          targetSheet.getRow(1).getCell(newColumnIndex).value = key;
          matchingRow.getCell(newColumnIndex).value = selectedRow[key];
          targetHeaders.push(key);
        }
      }
    });

    setUndoStack([...undoStack, { type: 'merge', originalRowValues, newRowValues: matchingRow.values, sheetIndex: targetSheetIndex }]);

    const updatedData = data.filter((_, i) => i !== currentRowIndex);
    setData(updatedData);
    workbook.worksheets[selectedSheet].spliceRows(currentRowIndex + 2, 1);
    setIsUndoDisabled(false);
  };

  const handleCellDoubleClick = (rowIndex, columnKey, value) => {
    setEditCell({ rowIndex, columnKey, value });
  };

  const handleEditChange = (event) => {
    setEditCell({ ...editCell, value: event.target.value });
  };

  const handleEditSave = () => {
    const updatedData = data.map((row, rowIndex) =>
      rowIndex === editCell.rowIndex ? { ...row, [editCell.columnKey]: editCell.value } : row
    );
    setUndoStack([...undoStack, { type: 'edit', data: data[editCell.rowIndex], rowIndex: editCell.rowIndex, sheetIndex: selectedSheet }]);
    setData(updatedData);
    const sheet = workbook.worksheets[selectedSheet];
    const row = sheet.getRow(editCell.rowIndex + 2);
    const columnIndex = sheet.getRow(1).values.indexOf(editCell.columnKey) + 1;
    row.getCell(columnIndex).value = editCell.value;
    setEditCell({ rowIndex: null, columnKey: null, value: '' });
    setIsUndoDisabled(false);
  };

  const handleNewRowSave = () => {
    const headers = workbook.worksheets[selectedSheet].getRow(1).values.slice(1);
    const newRowData = headers.reduce((acc, header) => {
      acc[header] = newRow[header] || '';
      return acc;
    }, {});
    setNewRow(newRowData);
    const updatedData = [...data, newRowData];
    setData(updatedData);
    const sheet = workbook.worksheets[selectedSheet];
    sheet.addRow(newRowData).commit();
    setNewRow({});
    setIsUndoDisabled(false);
  };

  const handleSaveFile = async () => {
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'output.xlsx';
    link.click();
  };

  const handlePushToDB = async () => {
    console.log('Push to DB functionality to be implemented');
  };

  const handleManualSelect = async () => {
    if (selectedManualRow) {
      await mergeRows(selectedManualRow.row, data[currentRowIndex], workbook.worksheets[targetSheetIndex].getRow(1).values.slice(1), workbook.worksheets[targetSheetIndex]);
      setManualSelectDialogOpen(false);
      setMergeDialogOpen(false);
      setIsUndoDisabled(false);
    }
  };

  const handleMergeAll = async () => {
    const targetSheet = workbook.worksheets[targetSheetIndex];
    if (!targetSheet) return;

    const targetHeaders = targetSheet.getRow(1).values.slice(1);
    const targetFirstNameIndex = targetHeaders.indexOf('firstName') + 1;
    const targetLastNameIndex = targetHeaders.indexOf('lastName') + 1;

    const rowsToDelete = [];

    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
      const selectedRow = data[rowIndex];
      let matchingRow;
      targetSheet.eachRow((row) => {
        const rowFirstName = row.getCell(targetFirstNameIndex).value?.toString().toLowerCase().trim();
        const rowLastName = row.getCell(targetLastNameIndex).value?.toString().toLowerCase().trim();
        const selectedFullNameParts = selectedRow.fullName.toLowerCase().split(' ');
        const selectedFirstName = selectedFullNameParts[0].trim();
        const selectedLastName = selectedFullNameParts[selectedFullNameParts.length - 1].trim();

        if (rowFirstName === selectedFirstName && rowLastName === selectedLastName) {
          matchingRow = row;
        }
      });

      if (matchingRow) {
        await mergeRows(matchingRow, selectedRow, targetHeaders, targetSheet);
        rowsToDelete.push(rowIndex);
      }
    }

    const updatedData = data.filter((_, index) => !rowsToDelete.includes(index));
    setData(updatedData);

    loadSheetData(workbook.worksheets[selectedSheet]);
    setIsUndoDisabled(false);
  };

  const handleUndo = () => {
    const lastAction = undoStack.pop();
    if (!lastAction) return;

    const sheet = workbook.worksheets[lastAction.sheetIndex];
    switch (lastAction.type) {
      case 'delete':
        const restoredData = [...data];
        restoredData.splice(lastAction.index, 0, lastAction.data);
        setData(restoredData);
        sheet.insertRow(lastAction.index + 2, lastAction.data);
        break;
      case 'edit':
        const editedData = [...data];
        editedData[lastAction.rowIndex] = lastAction.data;
        setData(editedData);
        const row = sheet.getRow(lastAction.rowIndex + 2);
        Object.entries(lastAction.data).forEach(([key, value]) => {
          const columnIndex = sheet.getRow(1).values.indexOf(key) + 1;
          row.getCell(columnIndex).value = value;
        });
        break;
      case 'merge':
        const { originalRowValues, newRowValues } = lastAction;
        const rowToUndo = sheet.getRow(Object.keys(originalRowValues)[0]);
        Object.entries(originalRowValues).forEach(([key, value]) => {
          rowToUndo.getCell(key).value = value;
        });
        const targetData = [...data, newRowValues];
        setData(targetData);
        break;
      default:
        break;
    }

    setUndoStack([...undoStack]);
    setIsUndoDisabled(true);
  };

  return (
    <Box sx={{ bgcolor: 'background.default', color: 'text.primary', minHeight: '100vh', p: 2 }} ref={tableRef}>
      <input type="file" onChange={handleFileUpload} />
      {workbook && (
        <>
          <Button onClick={handlePushToDB} variant="contained" sx={{ mt: 2, mr: 2 }}>Push to DB</Button>
          <Button onClick={handleSaveFile} variant="contained" sx={{ mt: 2, mr: 2 }}>Save File</Button>
          <Button onClick={handleMergeAll} variant="contained" sx={{ mt: 2, mr: 2 }}>Merge All</Button>
          <Button onClick={handleUndo} variant="contained" sx={{ mt: 2 }} disabled={isUndoDisabled}>Undo</Button>
          <Tabs value={selectedSheet} onChange={handleSheetChange} variant="scrollable" scrollButtons="auto">
            {workbook.worksheets.map((sheet, index) => (
              <Tab label={sheet.name} key={index} />
            ))}
          </Tabs>
          <Paper sx={{ mt: 2, height: '70vh', overflow: 'auto' }}>
            <Table stickyHeader>
              <TableHead>
                <TableRow>
                  {workbook.worksheets[selectedSheet].getRow(1).values.slice(1).map((header, index) => (
                    <TableCell key={index} sx={{ position: 'sticky', top: 0, zIndex: 3, backgroundColor: '#b8860b', color: 'black', fontWeight: 'bold' }}>{header}</TableCell>
                  ))}
                  <TableCell sx={{ position: 'sticky', top: 0, zIndex: 3, backgroundColor: '#b8860b', color: 'black', fontWeight: 'bold' }}>
                    <Tooltip title="Add New Row">
                      <IconButton onClick={handleNewRowSave}>
                        <AddIcon />
                      </IconButton>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {data.map((row, rowIndex) => (
                  <TableRow key={rowIndex}>
                    {Object.entries(row).map(([key, value], columnIndex) => (
                      <TableCell
                        key={columnIndex}
                        sx={
                          columnIndex < 2
                            ? {
                                position: 'sticky',
                                left: columnIndex === 0 ? 0 : 100,
                                zIndex: 2,
                                backgroundColor: '#b8860b',
                                color: 'black',
                                fontWeight: 'bold',
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis'
                              }
                            : {
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis'
                              }
                        }
                        onDoubleClick={() => handleCellDoubleClick(rowIndex, key, value)}
                      >
                        {editCell.rowIndex === rowIndex && editCell.columnKey === key ? (
                          <TextField
                            value={editCell.value}
                            onChange={handleEditChange}
                            onBlur={handleEditSave}
                            autoFocus
                          />
                        ) : typeof value === 'object' && value !== null && 'text' in value ? (
                          <a href={value.hyperlink} target="_blank" rel="noopener noreferrer" style={{ color: 'black', textDecoration: 'none', fontSize: 'small' }}>
                            {value.text}
                          </a>
                        ) : (
                          value
                        )}
                      </TableCell>
                    ))}
                    <TableCell
                      sx={{
                        position: 'sticky',
                        right: 0,
                        zIndex: 3,
                        backgroundColor: '#b8860b',
                        color: 'black',
                        fontWeight: 'bold',
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis'
                      }}
                    >
                      <Tooltip title="Delete">
                        <IconButton onClick={() => deleteRow(rowIndex)}>
                          <DeleteIcon />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="Merge">
                        <IconButton onClick={() => handleMergeClick(rowIndex)}>
                          <MergeIcon />
                        </IconButton>
                      </Tooltip>
                    </TableCell>
                  </TableRow>
                ))}
                <TableRow>
                  {workbook.worksheets[selectedSheet].getRow(1).values.slice(1).map((header, index) => (
                    <TableCell key={index}>
                      <TextField
                        placeholder={`Enter ${header}`}
                        value={newRow[header] || ''}
                        onChange={(e) => setNewRow({ ...newRow, [header]: e.target.value })}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleNewRowSave();
                        }}
                      />
                    </TableCell>
                  ))}
                  <TableCell>
                    <Tooltip title="Add New Row">
                      <IconButton onClick={handleNewRowSave}>
                        <AddIcon />
                      </IconButton>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              </TableBody>
            </Table>
          </Paper>
        </>
      )}

      <Dialog open={mergeDialogOpen} onClose={() => setMergeDialogOpen(false)}>
        <DialogTitle>Select Target Sheet</DialogTitle>
        <DialogContent>
          <Select
            value={targetSheetIndex !== null ? targetSheetIndex : ''}
            onChange={(e) => setTargetSheetIndex(e.target.value)}
            fullWidth
          >
            {workbook && workbook.worksheets.map((sheet, index) => (
              <MenuItem value={index} key={index}>{sheet.name}</MenuItem>
            ))}
          </Select>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setMergeDialogOpen(false)}>Cancel</Button>
          <Button onClick={handleMerge}>Merge</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={manualSelectDialogOpen} onClose={() => setManualSelectDialogOpen(false)}>
        <DialogTitle>Manual Row Selection</DialogTitle>
        <DialogContent>
          <Select
            value={selectedManualRow !== null ? selectedManualRow : ''}
            onChange={(e) => setSelectedManualRow(e.target.value)}
            fullWidth
          >
            {manualSelectOptions.map((option, index) => (
              <MenuItem value={option} key={index}>{`${option.firstName} ${option.lastName}`}</MenuItem>
            ))}
          </Select>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setManualSelectDialogOpen(false)}>Cancel</Button>
          <Button onClick={handleManualSelect}>Select</Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
};

export default ExcelHandler;



// C:\Users\Jude.Adenuga\kleandata\node_modules\electron\dist\electron.exe C:\Users\Jude.Adenuga\kleandata\public\main.js
// npx wait-on http://localhost:3000; npx electron