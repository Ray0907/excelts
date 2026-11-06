export const int = (rng, n) => Math.floor(rng() * n);
export const pick = (rng, a) => a[int(rng, a.length)];
