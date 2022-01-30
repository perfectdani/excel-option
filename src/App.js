import { useState, useRef } from 'react';
import { SpreadsheetComponent, SheetsDirective, SheetDirective, RangesDirective, RangeDirective, getRangeAddress } from '@syncfusion/ej2-react-spreadsheet';
import { Input, Select, Button, Modal, notification } from 'antd';
import { LineChartOutlined } from '@ant-design/icons'

import './App.css'

const { Option } = Select;

const colSTR = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

function App() {

  const ssRef = useRef();

  const [fileName, setFileName] = useState(null);
  const [columns, setColumns] = useState([]);
  const [data, setData] = useState([]);
  const [modalVisible, setModalVisible] = useState(false);
  const [applyPercent, setApplyPercent] = useState(null);
  const [applyColumns, setApplyColumns] = useState([]);
  const [applyType, setApplyType] = useState(null);

  const beforeOpen = (arg) => {
    setFileName(arg.file.name);
  }
  const beforeModal = () => {
    setModalVisible(true);
    let usedColIdx = ssRef.current.getActiveSheet().usedRange.colIndex;
    let colsArr = [];
    for (let col = 0; col < usedColIdx + 1; col++) {
      if (col < 26) {
        colsArr[col] = colSTR[col];
      } else {
        let a = parseInt(col / 26);
        let b = (col % 26);
        colsArr[col] = `${colSTR[a - 1]}${colSTR[b]}`;
      }
    }
    setColumns(colsArr);
  }

  function percentApply() {
    console.log(ssRef);
    if (applyType && applyPercent > 0 && applyColumns.length) {
      setModalVisible(false);
      setApplyType(null);
      setApplyPercent(null);
      setApplyColumns([]);
      let usedRowIdx = ssRef.current.getActiveSheet().usedRange.rowIndex;
      let usedColIdx = ssRef.current.getActiveSheet().usedRange.colIndex;
      let arr = [];
      let row = [];
      let index = 0;
      let current_row = 0; 
      ssRef.current.getData(getRangeAddress([0, 0, usedRowIdx, usedColIdx])).then((cells) => {
        cells.forEach((cell, key) => {
          if (index > usedColIdx) {
            arr = [...arr, row];
            row = [];
            index = 0;
            current_row++;
          }
          let newVlaue = cell.value;
          if (newVlaue &&newVlaue!==undefined&& applyColumns.includes(String(index))) {
            
            if (applyType === 1) {
              newVlaue = newVlaue * applyPercent / 100;
            } else if (applyType === 2) {
              newVlaue = newVlaue * (1 + applyPercent / 100);
            } else {
              newVlaue = newVlaue * (1 - applyPercent / 100);
            }
            newVlaue = Math.ceil(newVlaue);
            if(!isNaN(newVlaue))
              ssRef.current.sheets[0].rows[current_row].cells[index].value = newVlaue;
          }
          row = [...row, newVlaue];
          index++;
        })
        setData(arr);
        console.log(arr);
        // ssRef.current.sheets[0].ranges[0].dataSource = arr;
        
        
      }).then(()=>{
        ssRef.current.refresh(false)
        // ssRef.current.sheets[0].ranges[0].dataSource = arr;
      });
    } else {
      notification.warning({
        message: 'Please input correctly.'
      })
    }
  }

  const cancelModal = () => {
    setModalVisible(false);
    setApplyType(null);
    setApplyPercent(null);
    setApplyColumns([]);
  }

  const openCompleted = () => {
    ssRef.current.refresh(false)
  }
 
  
  return (
    <div className='App'>
      <Button type="text" icon={<LineChartOutlined />} id="modal-button" onClick={beforeModal} />
      <Modal
        style={{ top: 200 }}
        visible={modalVisible}
        onOk={percentApply}
        onCancel={cancelModal}
      >
        <label className='modal-label'>Type</label>
        <Select style={{ width: '100%' }} placeholder="Select Columns" value={applyType} onChange={(v) => setApplyType(v)}>
          <Option value={1}>Self</Option>
          <Option value={2}>Increase</Option>
          <Option value={3}>Decrease</Option>
        </Select>
        <label className='modal-label'>Percent</label>
        <Input type="number" suffix="%" min={0} placeholder='Input Value' value={applyPercent} onChange={(e) => setApplyPercent(e.target.value)} />
        <label className='modal-label'>Columns:</label>
        <Select mode="multiple" style={{ width: '100%' }} placeholder="Select Columns" value={applyColumns} onChange={(v) => setApplyColumns(v)}>
          {
            columns.map((column, index) => {
              return <Option key={index}>{column}</Option>;
            })
          }
        </Select>
      </Modal>
      <SpreadsheetComponent
        showSheetTabs={false}
        ref={ssRef}
        beforeOpen={beforeOpen}
        openCompleted={openCompleted}
        openUrl='https://ej2services.syncfusion.com/production/web-services/api/spreadsheet/open'
        saveUrl='https://ej2services.syncfusion.com/production/web-services/api/spreadsheet/save'
      >
        <SheetsDirective>
          <SheetDirective>
            <RangesDirective>
              <RangeDirective dataSource={data} showFieldAsHeader={false} />
            </RangesDirective>
          </SheetDirective>
        </SheetsDirective>
      </SpreadsheetComponent>
      <label className='filename'>{fileName}</label>
    </div>
  );
}

export default App
