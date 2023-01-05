// Copyright (c) 2017 PlanGrid, Inc.

import React, { Component } from 'react';
import XLSX from 'xlsx';

import CsvViewer from './csv-viewer';

class XlxsViewer extends Component {
  constructor(props) {
    super(props);
    this.state = this.parse();
  }

  parse() {
    const dataArr = new Uint8Array(this.props.data);
    const arr = [];


    for (let i = 0; i !== dataArr.length; i += 1) {
      arr.push(String.fromCharCode(dataArr[i]));
    }
    const workbook = XLSX.read(arr.join(''), { type: 'binary', cellStyles: true, cellDates: true, cellHTML: true });
    const names = Object.keys(workbook.Sheets);
    const sheets = names.map((name) => {
      let mc = -1;
      let mr = -1;
      if (workbook.Sheets[name]['!merges']) {
        workbook.Sheets[name]['!merges'].forEach((m) => {
          if (m.e.c > mc) {
            mc = m.e.c;
          }
          if (m.s.c > mc) {
            mc = m.s.c;
          }
          if (m.e.r > mr) {
            mr = m.e.r;
          }
          if (m.s.r > mr) {
            mr = m.s.r;
          }
        });
      }
      let j = XLSX.utils.sheet_to_json(workbook.Sheets[name], { blankRows: true, defval: null, header: 1 });
      if (j) {
        j.forEach((item, index) => {
          // const f = item.filter(value => index <= mr || (index > mr && value !== null));
          const f = item.filter(value => value !== null);
          if (f.length > 0) {
            const lastindex = item.lastIndexOf(f[f.length - 1]);
            if (lastindex > mc) {
              mc = lastindex;
            }
            if (index > mr) {
              mr = index;
            }
          }
        });
        if (mr > 10) {
          j = j.filter((item, index) => index <= mr);
        }
        j = j.map(item => item.slice(0, mc + 1));
        const ns = XLSX.utils.json_to_sheet(j, { skipHeader: true });
        return XLSX.utils.sheet_to_html(ns, { defval: null });
      }
      return XLSX.utils.sheet_to_html(workbook.Sheets[name], { defval: null });
    });
    return { sheets, names, curSheetIndex: 0 };
  }

  renderSheetNames(names) {
    const sheets = names.map((name, index) => (
      <input
        key={name}
        type="button"
        value={name}
        onClick={() => {
          this.setState({ curSheetIndex: index });
        }}
      />
    ));

    return (
      <div className="sheet-names">
        {sheets}
      </div>
    );
  }

  renderSheetData(sheet) {
    const csvProps = Object.assign({}, this.props, { data: sheet });
    return (
      <CsvViewer {...csvProps} />
    );
  }

  render() {
    const { sheets, names, curSheetIndex } = this.state;
    return (
      <div className="spreadsheet-viewer">
        {this.renderSheetNames(names)}
        <div style={{ background: 'white' }} dangerouslySetInnerHTML={{ __html: sheets[curSheetIndex || 0] }} />
      </div>
    );
  }
}

export default XlxsViewer;
