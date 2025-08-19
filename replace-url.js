
const { replaceInFile } = require('replace-in-file');

const path = require('path');



const env = process.argv[2]; // 'local' ou 'prod'



const url = env === 'local'

  ? 'https://localhost:3000/taskpane.html'

  : 'https://Eichenbergerr.github.io/EssaiAddin01/taskpane.html';



const manifestPath = path.resolve('manifest.xml');



replaceInFile({

  files: manifestPath,

  from: /<bt:Url\s+id="Taskpane\.Url"\s+DefaultValue="[^"]*"/,

  to: `<bt:Url id="Taskpane.Url" DefaultValue="${url}"`,

})

.then(results => {

  console.log('Remplacement effectué :', results);

  console.log('Fichier manifest modifié :', manifestPath);

})

.catch(error => {

  console.error('Erreur lors du remplacement :', error);

});

