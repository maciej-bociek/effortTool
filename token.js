for (const [key, value] of Object.entries(localStorage)) {
  if (
    key.includes('accesstoken') &&
    (key.includes('apihub.azure.com') || key.includes('service.flow.microsoft.com'))
  ) {
    try {
      const parsed = JSON.parse(value);
      if (parsed && parsed.secret && parsed.tokenType === 'Bearer') {
        console.log('Bearer', parsed.secret);
        break;
      }
    } catch {}
  }
}
