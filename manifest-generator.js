const getCurrentHostUrl = () => {
  const { protocol, hostname, port, pathname } = window.location;
  return `${protocol}//${hostname}${port ? `:${port}` : ''}${pathname}`;
};

const generateManifest = async (templateUrl, versionUrl) => {
  const responseManifestTemplate = await fetch(templateUrl);
  const manifestTemplate = await responseManifestTemplate.text();

  const responseVersion = await fetch(versionUrl);
  const versionContent = await responseVersion.text();
  const versionObject = JSON.parse(versionContent);

  const currentHostUrl = getCurrentHostUrl();

  // Extract the base URL dynamically up to "/office-addins"
  const baseRegex =
    /https:\/\/[\w.-]+\/static-server\/agent\/(beta|stable)\/office-addins/;
  const match = baseRegex.exec(currentHostUrl);
  const baseUrl = match[0];

  const manifestWithCurrentHost = manifestTemplate.replace(
    /https:\/\/localhost:4300/g,
    baseUrl
  );

  return manifestWithCurrentHost.replace(
    /<Version>.*?<\/Version>/,
    `<Version>${versionObject.version}</Version>`
  );
};
