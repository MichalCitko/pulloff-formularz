const cacheName = 'pulloff-shell-v2'
const appBase = self.registration.scope
const appShell = [
  appBase,
  new URL('manifest.webmanifest', appBase).href,
  new URL('pulloff-icon.svg', appBase).href,
  new URL('pulloff-192.png', appBase).href,
  new URL('pulloff-512.png', appBase).href,
]

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(cacheName).then((cache) => cache.addAll(appShell)))
  self.skipWaiting()
})

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) =>
      Promise.all(
        cacheNames
            .filter((name) => name.startsWith('pulloff-') && name !== cacheName)
          .map((name) => caches.delete(name)),
      ),
      ).then(() => self.clients.claim()),
  )
})

self.addEventListener('fetch', (event) => {
  const request = event.request
  const requestUrl = new URL(request.url)

  if (request.method !== 'GET' || requestUrl.origin !== self.location.origin) return

  event.respondWith(
    caches.match(request).then((cachedResponse) => {
      if (cachedResponse) return cachedResponse

      return fetch(request)
        .then(async (response) => {
          if (response.ok) {
            const cache = await caches.open(cacheName)
            await cache.put(request, response.clone())
          }
          return response
        })
        .catch(async () => (await caches.match(appBase)) ?? Response.error()),
    }),
  )
})