const CACHE_NAME = "smart-lecture-ai-v2";
const APP_SHELL = [
    "/",
    "/static/css/style.css",
    "/static/css/app-pages.css",
    "/static/manifest.json"
];

self.addEventListener("install", (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => cache.addAll(APP_SHELL))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener("activate", (event) => {
    event.waitUntil(
        caches.keys()
            .then((keys) => Promise.all(
                keys
                    .filter((key) => key !== CACHE_NAME)
                    .map((key) => caches.delete(key))
            ))
            .then(() => self.clients.claim())
    );
});

self.addEventListener("push", (event) => {
    let data = {};
    try {
        data = event.data ? event.data.json() : {};
    } catch (error) {
        data = {title: "Smart Lecture AI", body: event.data ? event.data.text() : "You have a new reminder."};
    }
    event.waitUntil(
        self.registration.showNotification(data.title || "Smart Lecture AI", {
            body: data.body || "You have a new study reminder.",
            icon: "/static/images/app-icon.png",
            badge: "/static/images/app-icon.png",
            data: {link: data.link || "/notifications"}
        })
    );
});

self.addEventListener("notificationclick", (event) => {
    event.notification.close();
    const link = event.notification.data && event.notification.data.link
        ? event.notification.data.link
        : "/notifications";
    event.waitUntil(clients.openWindow(link));
});

self.addEventListener("fetch", (event) => {
    const request = event.request;

    if (request.method !== "GET") {
        return;
    }

    const requestUrl = new URL(request.url);

    if (requestUrl.origin !== self.location.origin) {
        return;
    }

    if (request.mode === "navigate") {
        event.respondWith(
            fetch(request)
                .then((response) => {
                    const copy = response.clone();
                    caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
                    return response;
                })
                .catch(() => caches.match(request).then((cached) => cached || caches.match("/")))
        );
        return;
    }

    event.respondWith(
        caches.match(request).then((cached) => cached || fetch(request).then((response) => {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
            return response;
        }))
    );
});