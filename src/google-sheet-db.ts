/* tslint:disable: member-ordering */
import * as uuid from 'uuid/v4';
import GoogleSheet, { IOptions } from './google-sheet';
import { ISheetProperty } from './types';

interface IRecord {
  _index?: number;
  _id?: string;
  [key: string]: any;
}

interface IColumnMap {
  [key: string]: {
    index: number;
  };
}

interface ICollection<T extends IRecord = IRecord> {
  name: string;
  sheet: ISheetProperty;
  columns: string[];
  column_map: IColumnMap;
  data: T[];
  initialized?: boolean;
}

const defaultPredicateFn = (): boolean => true;

const getColumnByIndex = (index: number): string => {
  if (index <= 0) {
    throw new Error('Index out range');
  }

  // tslint:disable-next-line: max-line-length
  const AZ = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

  if (index <= 26) {
    return AZ[index - 1];
  } else {
    const a = (index - 1) % 26;
    const b = Math.floor((index - 1) / 26) - 1;

    return AZ[b] + AZ[a];
  }
};

class GoogleSheetDb {
  private googleSheet: GoogleSheet;
  private collections: ICollection[] = [];

  constructor(spreadsheetId: string, options?: IOptions) {
    this.googleSheet = new GoogleSheet(spreadsheetId, options);
  }

  public async initialize() {
    await this.googleSheet.authenticate();
    const sheets = await this.googleSheet.getSheets();

    this.collections.length = 0;

    for (const sheet of sheets) {
      this.collections.push({
        name: sheet.title,
        sheet,
        column_map: {},
        columns: [],
        data: [],
        initialized: false,
      });
    }
  }

  private async createSheet(collectionName: string) {
    const sheets = await this.googleSheet.createSheet(collectionName);
    const sheet = sheets.find(x => x.title === collectionName);
    await this.googleSheet.applyHeaderStyle(sheet.sheetId);
    return sheet;
  }

  private async syncHeader(collection: ICollection) {
    const sheetName = collection.name;
    const rangeStart = 'A1';
    const rangeEnd = `${getColumnByIndex(collection.columns.length)}1`;
    const values = [collection.columns.map(x => {
      return x && x.startsWith('_column_') ? '' : x;
    })];
    await this.googleSheet.writeData(`${sheetName}!${rangeStart}:${rangeEnd}`, values);
  }

  private async getCollectionDataBySheet(sheet: ISheetProperty): Promise<ICollection> {
    const data = [];
    const columns = [];
    const columnMap: IColumnMap = {};

    if (sheet.gridProperties.rowCount && sheet.gridProperties.columnCount) {
      const sheetName = sheet.title;
      const rangeStart = `${getColumnByIndex(1)}1`;
      const rangeEnd = `${getColumnByIndex(sheet.gridProperties.columnCount)}${sheet.gridProperties.rowCount}`;
      const rows = await this.googleSheet.readData(`${sheetName}!${rangeStart}:${rangeEnd}`);

      const headerRow = rows[0];
      headerRow.forEach((x, i) => {
        if (!x) {
          columns.push(`_column_${i}`);
          columnMap[`_column_${i}`] = { index: i };
        } else {
          columns.push(x);
          columnMap[x] = { index: i };
        }
      });

      rows.splice(0, 1);

      rows.forEach((x, index) => {
        const item: IRecord = {
          _index: index,
          _id: uuid(),
        };
        // tslint:disable-next-line: prefer-for-of
        for (let i = 0; i < columns.length; i++) {
          item[columns[i]] = x[columnMap[columns[i]].index];
        }

        data.push(item);
      });
    }

    return {
      name: sheet.title,
      sheet,
      columns,
      column_map: columnMap,
      data,
      initialized: true,
    }
  }

  private async initCollection(collection: ICollection) {
    if (collection.initialized) {
      // already initialized
      return;
    }

    
    const sheet = collection.sheet;
    const _collection = await this.getCollectionDataBySheet(sheet);

    collection.columns = _collection.columns;
    collection.column_map = _collection.column_map;
    collection.data = _collection.data;
    collection.initialized = true;
  }

  public async insert<T extends IRecord>(collectionName: string, record: T) {
    await this.insertMany(collectionName, [record]);
  }

  public async insertMany<T extends IRecord>(collectionName: string, records: T[]) {
    if (!records.length) {
      return 0;
    }

    let collection = this.collections.find(x => x.name === collectionName);

    if (!collection) {
      collection = {
        columns: [],
        column_map: {},
        data: [],
        name: collectionName,
        sheet: null,
      };
    }

    if (!collection.sheet) {
      collection.sheet = await this.createSheet(collection.name);
    } else {
      await this.initCollection(collection);
    }

    const values = [];
    let syncHeader = false;
    records.forEach(x => {
      const rColumns = Object.keys(x).filter(y => y !== '_index' && y !== '_id');

      const value = [];

      rColumns.forEach(y => {
        if (!collection.column_map[y]) {
          collection.column_map[y] = {
            index: collection.columns.length,
          };
          collection.columns.push(y);
          syncHeader = true;
        }

        value[collection.column_map[y].index] = x[y];
      });

      values.push(value);
    });

    const sheetName = collection.name;
    const rangeStart = `A${collection.data.length + 2}`;
    const rangeEnd = `${getColumnByIndex(collection.columns.length)}${values.length + collection.data.length + 1}`;

    if (syncHeader) {
      await this.syncHeader(collection);
    }
    await this.googleSheet.writeData(`${sheetName}!${rangeStart}:${rangeEnd}`, values);

    records.forEach((item, i) => {
      item._index = collection.data.length + i;
      item._id = uuid();
    });

    Array.prototype.push.apply(collection.data, records);

    return values.length;
  }

  public async find<T extends IRecord>(
    collectionName: string,
    predicate?: (value: T, index: number, obj: T[]) => boolean,
  ) {
    const collection = this.collections.find(x => x.name === collectionName);

    if (!collection) {
      return [];
    }

    await this.initCollection(collection);

    return collection.data.filter(predicate || defaultPredicateFn);
  }

  public async findOne<T extends IRecord>(
    collectionName: string,
    predicate?: (value: T, index: number, obj: T[]) => boolean,
  ) {
    return (await this.find(collectionName, predicate))[0];
  }

  public async delete<T extends IRecord>(
    collectionName: string,
    predicate: (value: T, index: number, obj: T[]) => boolean,
  ) {
    const collection = this.collections.find(x => x.name === collectionName);

    if (!collection) {
      return 0;
    }

    await this.initCollection(collection);

    const deleteIndex = [];
    collection.data.forEach((value: T, index, obj: T[]) => {
      if (predicate(value, index, obj)) {
        deleteIndex.push(index);
      }
    });

    for (let i = deleteIndex.length - 1; i >= 0; i--) {
      const index = deleteIndex[i];
      await this.googleSheet.removeRows(collection.sheet.sheetId, index + 1, 1);
      this.collections.splice(index, 1);
    }

    collection.data.forEach((item, i) => {
      item._index = i;
      item._id = uuid();
    });

    return deleteIndex.length;
  }

  public async update<T extends IRecord>(collectionName: string, record: T) {
    const collection = this.collections.find(x => x.name === collectionName);

    if (!collection) {
      throw new Error('Collection not exists');
    }

    if (!record) {
      throw new Error('Invalid record');
    }

    if (typeof record._index !== 'number') {
      throw new Error('Invalid index');
    }

    await this.initCollection(collection);

    const values = [];
    let syncHeader = false;
    const rColumns = Object.keys(record).filter(y => y !== '_index' && y !== '_id');

    const value = [];

    rColumns.forEach(y => {
      if (!collection.column_map[y]) {
        collection.column_map[y] = {
          index: collection.columns.length,
        };
        collection.columns.push(y);
        syncHeader = true;
      }

      value[collection.column_map[y].index] = record[y];
    });

    values.push(value);

    const sheetName = collection.name;
    const rangeStart = `A${record._index + 2}`;
    const rangeEnd = `${getColumnByIndex(collection.columns.length)}${values.length + record._index + 1}`;

    if (syncHeader) {
      await this.syncHeader(collection);
    }
    await this.googleSheet.writeData(`${sheetName}!${rangeStart}:${rangeEnd}`, values);

    collection.data[record._index] = record;

    return true;
  }

  public async refreshRecord<T extends IRecord>(collectionName: string, record: T) {
    const collection = this.collections.find(x => x.name === collectionName);

    if (!collection) {
      throw new Error('Collection not exists');
    }

    if (!record) {
      throw new Error('Invalid record');
    }

    if (typeof record._index !== 'number') {
      throw new Error('Invalid index');
    }

    await this.initCollection(collection);

    const sheetName = collection.name;
    const rangeStart = `A${record._index + 2}`;
    const rangeEnd = `${getColumnByIndex(collection.columns.length)}${1 + record._index + 1}`;

    const rows = await this.googleSheet.readData(`${sheetName}!${rangeStart}:${rangeEnd}`);

    const item: IRecord = {
      _index: record._index,
      _id: record._id,
    };

    const columns = collection.columns;
    const columnMap = collection.column_map;
    const rowItem = rows[0];
    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < columns.length; i++) {
      item[columns[i]] = rowItem[columnMap[columns[i]].index];
    }

    Object.keys(record).forEach(key => {
      if (key === '_index' || key === '_id') {
        return;
      }

      (record as any)[key] = item[key];
    })
  }

  public async refreshCollection(collectionName: string) {
    const collection = this.collections.find(x => x.name === collectionName);

    if (!collection) {
      throw new Error('Collection not exists');
    }

    const sheets = await this.googleSheet.getSheets();

    const sheet = sheets.find(x => x.title === collectionName);

    if (!sheet) {
      throw new Error('Sheet not found');
    }

    const _collection = await this.getCollectionDataBySheet(sheet);

    collection.name = sheet.title;
    collection.sheet = sheet;
    collection.columns = _collection.columns;
    collection.column_map = _collection.column_map;
    collection.data = _collection.data;
    collection.initialized = true;
  }

  public async refreshAll() {
    const sheets = await this.googleSheet.getSheets();

    this.collections.length = 0;

    for (const sheet of sheets) {
      const _collection = await this.getCollectionDataBySheet(sheet);

      this.collections.push({
        name: sheet.title,
        sheet,
        columns: _collection.columns,
        column_map: _collection.column_map,
        data: _collection.data,
        initialized: true,
      });
    }
  }
}

export default GoogleSheetDb;
