// Startup checks are read-only and safe to retry after transient failures.
export async function connect(api, {
  sleep = ms => new Promise(resolve => setTimeout(resolve, ms)),
  log = message => console.error(message),
} = {}) {
  let delay = 2000;
  while (true) {
    try {
      const me = await api('getMe', {});
      const webhook = await api('getWebhookInfo', {});
      return {me, webhook};
    } catch (error) {
      if (error.code === 401 || error.code === 404 || error.code === 409) {
        throw Error(`Telegram startup rejected (${error.code}). Check the token or other bot instances.`);
      }
      const wait = Math.max(delay, (error.retryAfter || 0) * 1000);
      log(`Telegram startup connection failed. Retrying in ${Math.ceil(wait / 1000)} seconds.`);
      await sleep(wait);
      delay = Math.min(delay * 2, 60000);
    }
  }
}
