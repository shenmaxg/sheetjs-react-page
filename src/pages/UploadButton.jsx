import React from 'react';
import { Upload, Button } from 'antd';
import { UploadOutlined } from '@ant-design/icons';
import { blob_to_wb } from './Excel2AntUtil';

export default ({ setSheet, children }) => {
  const onUploadExcel = async ({ file }) => {
    if (file.status === 'done') {
      const { originFileObj } = file;
      const workbook = await blob_to_wb(originFileObj);

      let sheet;
      for (const sheetName in workbook.Sheets) {
        if (workbook.Sheets.hasOwnProperty(sheetName)) {
          sheet = workbook.Sheets[sheetName];
          break;
        }
      }

      setSheet(sheet);
    }
  };

  return (
    <Upload
      onChange={onUploadExcel}
      accept="application/vnd.ms-excel,
                        application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    >
      <Button>
        <UploadOutlined />
        {children}
      </Button>
    </Upload>
  );
};
