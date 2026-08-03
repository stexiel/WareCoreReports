const selfsigned = require('selfsigned');
const fs = require('fs');
const path = require('path');

async function generateCerts() {
  const attrs = [{ name: 'commonName', value: 'localhost' }];
  const pems = await selfsigned.generate(attrs, { days: 365 });

  if (pems.private && pems.cert) {
    fs.writeFileSync(path.join(__dirname, 'certs', 'key.pem'), pems.private);
    fs.writeFileSync(path.join(__dirname, 'certs', 'cert.pem'), pems.cert);
    console.log('SSL сертификаты сгенерированы в папке certs/');
    console.log('key.pem и cert.pem');
  } else {
    console.error('Ошибка генерации сертификатов');
    console.log('PEMs:', pems);
  }
}

generateCerts();
