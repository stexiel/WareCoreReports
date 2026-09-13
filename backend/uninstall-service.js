var Service = require('node-windows').Service;

var svc = new Service({
  name: 'WareCore Backend',
  description: 'WareCore Reports Backend Server',
  script: 'C:\\Users\\Administrator\\Desktop\\асер\\WareCore Reports\\backend\\server.js',
  nodeOptions: ['--max-old-space-size=4096']
});

svc.on('uninstall', function(){
  console.log('Service uninstalled');
});

svc.uninstall();
