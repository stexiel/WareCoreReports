var Service = require('node-windows').Service;

var svc = new Service({
  name: 'WareCore Frontend',
  script: 'C:\\Users\\Administrator\\Desktop\\асер\\WareCore Reports\\frontend-react\\node_modules\\vite\\bin\\vite.js'
});

svc.on('uninstall', function () {
  console.log('Service uninstalled');
});

svc.uninstall();
