var Service = require('node-windows').Service;

var svc = new Service({
  name: 'WareCore Backend',
  description: 'WareCore Reports Backend Server',
  script: 'C:\\Users\\Administrator\\Desktop\\асер\\WareCore Reports\\backend\\server.js',
  nodeOptions: ['--max-old-space-size=4096']
});

svc.on('install', function(){
  svc.start();
  console.log('Service installed and started');
});

svc.on('start', function(){
  console.log('Service started');
});

svc.on('stop', function(){
  console.log('Service stopped');
});

svc.install();
