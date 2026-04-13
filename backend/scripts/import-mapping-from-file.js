const fs = require('fs').promises;
const path = require('path');

const MAPPING_FILE = path.join(__dirname, '../data/mapping-emplacement.json');
const INPUT_FILE = '/tmp/mapping.txt';

async function importMapping() {
  console.log('🚀 Import du mapping depuis fichier...');
  
  const data = await fs.readFile(INPUT_FILE, 'utf8');
  const lines = data.trim().split('\n');
  const mapping = {};
  
  let imported = 0;
  let errors = 0;
  
  lines.forEach((line, index) => {
    const parts = line.split('\t');
    
    if (parts.length >= 5) {
      const code = parts[0].trim();
      const description = parts[1].trim();
      const type = parts[2].trim();
      const typeEmplacement = parts[3].trim();
      const detail = parts[4].trim();
      
      mapping[code] = {
        description,
        type,
        typeEmplacement,
        detail
      };
      
      imported++;
    } else {
      console.warn(`⚠️  Ligne ${index + 1} ignorée (format invalide)`);
      errors++;
    }
  });
  
  console.log(`📦 ${imported} emplacements importés`);
  console.log(`❌ ${errors} lignes ignorées`);
  
  await fs.writeFile(MAPPING_FILE, JSON.stringify(mapping, null, 2));
  console.log('✅ Mapping sauvegardé dans', MAPPING_FILE);
  
  // Afficher un échantillon
  const keys = Object.keys(mapping).slice(0, 3);
  console.log('\n📋 Échantillon:');
  keys.forEach(key => {
    console.log(`  ${key}: ${mapping[key].description}`);
  });
}

importMapping().catch(console.error);
