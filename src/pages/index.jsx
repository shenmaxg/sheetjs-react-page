import React from 'react';
import styles from './index.less';
import UploadButton from './UploadButton';
import { Table } from 'antd';
import {
  sheet_to_ant_columns,
  sheet_to_ant_dataSource,
  ant_table_to_json,
} from './Excel2AntUtil';

const sheetToTable = function(sheet) {
  if (sheet) {
    const columns = sheet_to_ant_columns(sheet);
    const dataSource = sheet_to_ant_dataSource(sheet, columns);

    return { columns, dataSource };
  }

  return {};
};

export default () => {
  const [sheet, setSheet] = React.useState();
  const { columns, dataSource } = sheetToTable(sheet);
  const json = ant_table_to_json(columns, dataSource);

  return (
    <div className={styles.main}>
      <div>
        <UploadButton setSheet={setSheet}>点击上传 Excel</UploadButton>
      </div>
      {sheet && (
        <Table
          columns={columns}
          dataSource={dataSource}
          bordered
          pagination={false}
          size="small"
        />
      )}
      {json && (
        <pre style={{ paddingTop: '20px' }}>
          {JSON.stringify(json, null, 2)}
        </pre>
      )}
    </div>
  );
};
