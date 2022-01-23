import React, { useContext, useState, useEffect, useRef } from 'react';
import { ExcelRenderer } from 'react-excel-renderer'
import { ExportJsonCsv } from 'react-export-json-csv';
import { Table, Input, Select, Button, Popconfirm, Form, Modal } from 'antd';
import { PercentageOutlined, StockOutlined } from '@ant-design/icons'
import './App.css'

const { Option } = Select;

const EditableContext = React.createContext(null);

const EditableRow = ({ index, ...props }) => {
  const [form] = Form.useForm();
  return (
    <Form form={form} component={false}>
      <EditableContext.Provider value={form}>
        <tr {...props} />
      </EditableContext.Provider>
    </Form>
  );
};

const EditableCell = ({
  title,
  editable,
  children,
  dataIndex,
  record,
  handleSave,
  ...restProps
}) => {
  const [editing, setEditing] = useState(false);
  const inputRef = useRef(null);
  const form = useContext(EditableContext);
  useEffect(() => {
    if (editing) {
      inputRef.current.focus();
    }
  }, [editing]);

  const toggleEdit = () => {
    setEditing(!editing);
    form.setFieldsValue({
      [dataIndex]: record[dataIndex],
    });
  };

  const save = async () => {
    try {
      const values = await form.validateFields();
      toggleEdit();
      handleSave({ ...record, ...values });
    } catch (errInfo) {
      console.log('Save failed:', errInfo);
    }
  };

  let childNode = children;

  if (editable) {
    childNode = editing ? (
      <Form.Item
        style={{
          margin: 0,
        }}
        name={dataIndex}
        rules={[
          {
            required: true,
            message: `${title} is required.`,
          },
        ]}
      >
        <Input ref={inputRef} onPressEnter={save} onBlur={save} />
      </Form.Item>
    ) : (
      <div
        className="editable-cell-value-wrap"
        style={{
          paddingRight: 24,
        }}
        onClick={toggleEdit}
      >
        {children}
      </div>
    );
  }

  return <td {...restProps}>{childNode}</td>;
};

class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      fileName: null,
      headers: [],
      columns: [],
      dataSource: [],
      modal1Visible: false,
      percent: null,
      applyColumns: [],
      talbleHeight: window.innerHeight - 205
    };
  }

  handleSave = (row) => {
    const newData = [...this.state.dataSource];
    const index = newData.findIndex((item) => row.id === item.id);
    const item = newData[index];
    newData.splice(index, 1, { ...item, ...row });
    this.setState({
      dataSource: newData,
    });
  };


  fileHandler = (event) => {
    let fileObj = event.target.files[0];
    ExcelRenderer(fileObj, (err, resp) => {
      if (err) {
        console.log(err);
      } else {
        let columnsArr = [];
        resp.cols.map(col => {
          columnsArr = [...columnsArr, {
            dataIndex: col.key,
            title: col.name,
            onCell: (record) => ({
              record,
              editable: true,
              dataIndex: col.key,
              title: col.name,
              handleSave: this.handleSave,
            })
          }]
        });
        let rowsArr = [];
        resp.rows.map((row, index) => {
          rowsArr = [...rowsArr, {
            ...row,
            id: index
          }];
        });
        this.setState({
          fileName: fileObj['name'],
          headers: resp.cols,
          columns: columnsArr,
          dataSource: rowsArr
        });
      }
    });
  }

  percentApply = () => {
    let newData = [...this.state.dataSource];
    this.state.dataSource.map((row, index) => {
      this.state.applyColumns.map(col => {
        newData[index][col] = row[col] * this.state.percent / 100;
      })
    })
    this.setState({
      dataSource: newData
    });
  }

  render() {
    const { dataSource, columns, headers } = this.state;
    const components = {
      body: {
        row: EditableRow,
        cell: EditableCell,
      },
    };
    const handleResize = () => {
      this.setState({talbleHeight: window.innerHeight - 205});
    }
    window.addEventListener("resize", handleResize);
    return (
      <div className='App'>
        <div className='action'>
          <input type="file" accept=".csv, text/csv" onChange={this.fileHandler} />
          {
            this.state.columns.length ?
              <React.Fragment>
                <Button shape="circle" icon={<StockOutlined />} onClick={() => this.setState({ modal1Visible: true })} style={{ marginLeft: '50px' }} />
                <Modal
                  style={{ top: 50 }}
                  visible={this.state.modal1Visible}
                  onOk={this.percentApply}
                  onCancel={() => this.setState({ modal1Visible: false })}
                >
                  <label className='modal-label'>Percent</label>
                  <Input type="number" suffix="%" min={0} onChange={(e) => this.setState({ percent: e.target.value })} />
                  <label className='modal-label'>Columns:</label>
                  <Select mode="multiple" style={{ width: '100%' }} placeholder="Select Columns" onChange={(v) => this.setState({ applyColumns: v })}>
                    {
                      this.state.columns.map((column) => {
                        return <Option key={column.dataIndex}>{column.title}</Option>;
                      })
                    }
                  </Select>
                </Modal>
                <ExportJsonCsv headers={headers} items={dataSource} style={{ marginLeft: '50px' }}>Export</ExportJsonCsv>
              </React.Fragment>
              : null
          }
        </div>
        <Table
          bordered
          rowKey="id"
          components={components}
          columns={columns}
          dataSource={dataSource}
          pagination={{
            position: ['bottomCenter']
          }}
          rowClassName={() => 'editable-row'}
          scroll={{ y: this.state.talbleHeight }}
        />
      </div>
    );
  }
}

export default App
