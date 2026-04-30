  execSync       node:child_process
  existsSync, rmSync       node:fs
  dirname, resolve       node:path
  fileURLToPath       node:url

      __dirname   dirname(fileURLToPath(.     .meta.url));
      repoRoot   resolve(__dirname, '..')
      docsRoot   resolve(repoRoot, 'docs')


    lastTag
    
  lastTag   execSync('git describe --tags --abbrev=0', 
    encoding: 'utf-8',
    cwd: repoRoot
   ).trim()

      
  console.log('No git tags found. Skipping stable version preparation.')
  process.exit(0)


console.log(`Creating stable version from tag: ${lastTag}`)


    (.     p of ['versioned_docs', 'versioned_sidebars', 'versios.json']) 
        fullPath = resolve(docsRoot, p);
     (existsSync(fullPath)) 
    rmSync(fullPath,   recursive: true, force: true });
  


    
  rmSync(resolve(docsRoot, 'docs'),  recursive:  
  execSync(`git restore --source="${lastTag}" -- docs/docs/ docs/src/config/sidebars.ts`, 
    cwd: repoRoot
  

  
  execSync(`npx docusaurus docs:version "${lastTag}"`, 
    cwd: docsRoot,
    stdio: 'inherit'
  

  console.log(`Stable version created successfully from tag ${lastTag}`);


  
  execSync('git restore -- docs/docs/ docs/src/config/sidebars.ts', 
    cwd: repoRoot
  
  
  execSync('git clean -fd docs/docs/', 
    cwd: repoRoot


