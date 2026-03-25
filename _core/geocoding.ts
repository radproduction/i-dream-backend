import { ENV } from "./env";

type CacheEntry = {
  address: string;
  expiresAt: number;
};

const CACHE_TTL_MS = 1000 * 60 * 60; // 1 hour
const cache = new Map<string, CacheEntry>();

function cacheKey(lat: number, lng: number) {
  return `en:${lat.toFixed(5)},${lng.toFixed(5)}`;
}

export async function reverseGeocode(lat: number, lng: number) {
  const key = cacheKey(lat, lng);
  const cached = cache.get(key);
  if (cached && cached.expiresAt > Date.now()) {
    return cached.address;
  }

  const url = new URL("https://nominatim.openstreetmap.org/reverse");
  url.searchParams.set("format", "json");
  url.searchParams.set("lat", String(lat));
  url.searchParams.set("lon", String(lng));
  url.searchParams.set("zoom", "18");
  url.searchParams.set("addressdetails", "1");
  url.searchParams.set("accept-language", "en");

  try {
    const response = await fetch(url.toString(), {
      headers: {
        accept: "application/json",
        "accept-language": "en",
        "user-agent":
          ENV.geocodeUserAgent || "i-dream-hrms/1.0 (contact: admin)",
      },
    });

    if (!response.ok) return null;
    const data = (await response.json()) as { display_name?: string };
    const address = data?.display_name?.trim() || null;
    if (address) {
      cache.set(key, { address, expiresAt: Date.now() + CACHE_TTL_MS });
    }
    return address;
  } catch {
    return null;
  }
}
