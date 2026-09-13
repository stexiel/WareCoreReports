var Service = require('node-windows').Service;

var svc = new Service({
  name: 'WareCore Frontend',
  description: 'WareCore Reports Frontend (React + Vite)',
  script: 'C:\\Users\\Administrator\\Desktop\\асер\\WareCore Reports\\frontend-react\\node_modules\\vite\\bin\\vite.js',
  scriptOptions: '--host 0.0.0.0 --port 5173',
  nodeOptions: ['--max-old-space-size=1024']
});

svc.on('install', function () {
  svc.start();
  console.log('Service installed and started');
});

svc.on('alreadyinstalled', function () {
  console.log('Service is already installed');
});

svc.on('start', function () {
  console.log('Service started');
});

svc.on('stop', function () {
  console.log('Service stopped');
});

svc.install();
