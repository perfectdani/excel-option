import { useState, useRef } from 'react';
import { SpreadsheetComponent, SheetsDirective, SheetDirective, RangesDirective, RangeDirective, getRangeAddress } from '@syncfusion/ej2-react-spreadsheet';
import { Input, Select, Button, Modal, notification, Row, Col } from 'antd';
import { LineChartOutlined,SaveOutlined } from '@ant-design/icons'

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
  
  const save = (ext) =>{
    let usedRowIdx = ssRef.current.getActiveSheet().usedRange.rowIndex;
    let usedColIdx = ssRef.current.getActiveSheet().usedRange.colIndex;
    let index = 0;
    let dollarUSLocale = Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
    
      // These options are needed to round to whole numbers if that's what you want.
      //minimumFractionDigits: 0, // (this suffices for whole numbers, but will print 2500.10 as $2,500.1)
      //maximumFractionDigits: 0, // (causes 2500.99 to be printed as $2,501)
    });
    ssRef.current.getData(getRangeAddress([0, 0, usedRowIdx, usedColIdx])).then((cells) => {
      console.log(cells);
      let txtData = "";
      cells.forEach((cell,key)=>{
        if (index > usedColIdx) {
          index = 0;
          txtData += "\r\n";
        }
        let newValue = cell.value;
        console.log(cell.format);
        
        if(cell.format!==undefined){
          // newValue = dollarUSLocale.format(newValue);
          newValue = '$'+newValue;
          console.log(newValue);
        }

        if(newValue===undefined)
          txtData += ",";
        else{
          txtData += `"`+newValue+`"`+",";
        }
          
        index++;
      });
      const filename = "sample."+ext;
      download(filename,txtData);
    }).then(()=>{
      
      // ssRef.current.sheets[0].ranges[0].dataSource = arr;
    });
  }
  async function download(filename, text) {
    const image = new Blob([text], { type: 'application/octet-binary' })
    const opts = {
      types: [{
        description: 'CSV file',
        accept: {'text/csv': ['.csv']},
      }],
    };
    if( window.showSaveFilePicker ) {
      const handle = await window.showSaveFilePicker(opts);
      const writable = await handle.createWritable();
      await writable.write( image );
      writable.close();
    }
    else {
      const saveImg = document.createElement( "a" );
      saveImg.href = URL.createObjectURL( image );
      saveImg.download= "image.png";
      saveImg.click();
      setTimeout(() => URL.revokeObjectURL( saveImg.href ), 60000 );
    }
  }
  
  return (
    <div className='App'>
      <Button type="text" icon={<LineChartOutlined />} id="modal-button" onClick={beforeModal} />
      <Button type="text" icon={<SaveOutlined />} id="save_csv" onClick={()=>save('csv')} />
      <Modal
        style={{ top: 200 }}
        visible={modalVisible}
        onOk={percentApply}
        onCancel={cancelModal}
        footer={null}
      >
        <label className='modal-label'>Type</label>
        <Select style={{ width: '100%' }} placeholder="Select Columns" value={applyType} onChange={(v) => setApplyType(v)}>
          <Option value={1}>Self</Option>
          <Option value={2}>Increase</Option>
          <Option value={3}>Decrease</Option>
        </Select>
        <label className='modal-label'>Percent</label>
        <Input type="number" suffix="%" min={0} placeholder='Input Value' value={applyPercent} onChange={(e) => setApplyPercent(e.target.value)} />
        <Row>
          <Col md={15} sm={12} xs={12}>
            <label className='modal-label'>Columns:</label>
            <Select mode="multiple" style={{ width: '100%' }} placeholder="Increase Type" value={applyColumns} onChange={(v) => setApplyColumns(v)}>
              {
                columns.map((column, index) => {
                  return <Option key={index}>{column}</Option>;
                })
              }
            </Select>
          </Col>
          <Col md={9} sm={12} xs={12} style={{ textAlign: 'right', paddingTop: 62 }}>
            <Button onClick={cancelModal} type="danger" >Cancel</Button>
            <Button onClick={percentApply} type="primary" style={{ marginLeft: 12 }}>Ok</Button>
          </Col>
        </Row>
      </Modal>
      <SpreadsheetComponent
        showSheetTabs={false}
        ref={ssRef}
        beforeOpen={beforeOpen}
        allowSave= {false}
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
