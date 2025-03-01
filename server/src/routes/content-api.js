export default [
  {
    method: 'GET',
    path: '/',
    handler: 'controller.index',
    config: {
      policies: [],
    },
  },
  {
    method: 'GET',
    path: '/dropdown-values',
    handler: 'controller.getDropDownData',
    config: {
      policies: [],
      auth: false,
    },
  },
  {
    method: 'GET',
    path: '/table-data',
    handler: 'controller.getTableData',
    config: {
      policies: [],
      auth: false,
    },
  },
  {
    method: 'GET',
    path: '/download-excel',
    handler: 'controller.downloadExcel',
    config: {
      policies: [],
      auth: false,
    },
  },
];
