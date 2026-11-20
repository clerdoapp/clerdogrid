<?php
declare(strict_types=1);

namespace Grid;
/**
 * Addresses of Clerdo's own servers (*.clerdo.app). This host's DNS is often slow, and a lookup that times out
 * made sign-in and app-to-app calls fail or run past PHP's time limit (a 500). The Clerdo sites share this
 * server, so a call first goes to this server's own address; an address that worked is remembered for a day.
 * If it doesn't answer (or its certificate doesn't match), the caller unpins and does a normal lookup.
 */
final class DnsCache
{
    private const TTL = 86400;

    /** curl options that use the remembered address for $url's host, or [] when there is none. */
    public static function options(string $url): array
    {
        $hp = self::hostPort($url);
        if ($hp === null) {
            return [];
        }
        $ip = self::load()[$hp[0]] ?? self::ownAddress();
        return is_string($ip) ? [CURLOPT_RESOLVE => [$hp[0] . ':' . $hp[1] . ':' . (str_contains($ip, ':') ? '[' . $ip . ']' : $ip)]] : [];
    }

    /** Undoes options() on a reused handle, so the next attempt looks the host up again. */
    public static function unpin(\CurlHandle $ch, string $url): void
    {
        $hp = self::hostPort($url);
        if ($hp !== null) {
            curl_setopt($ch, CURLOPT_RESOLVE, ['-' . $hp[0] . ':' . $hp[1]]);
            self::save(array_diff_key(self::load(), [$hp[0] => 1]));
        }
    }

    /** After a successful request: remember the address it connected to. */
    public static function remember(\CurlHandle $ch, string $url): void
    {
        $hp = self::hostPort($url);
        $ip = (string)curl_getinfo($ch, CURLINFO_PRIMARY_IP);
        if ($hp === null || filter_var($ip, FILTER_VALIDATE_IP) === false) {
            return;
        }
        $all = self::load();
        if (($all[$hp[0]] ?? null) !== $ip) {
            $all[$hp[0]] = $ip;
            self::save($all);
        }
    }

    /** This server's public address, as the web server reports it. */
    private static function ownAddress(): ?string
    {
        $ip = (string)($_SERVER['SERVER_ADDR'] ?? '');
        return filter_var($ip, FILTER_VALIDATE_IP, FILTER_FLAG_NO_PRIV_RANGE | FILTER_FLAG_NO_RES_RANGE) !== false ? $ip : null;
    }

    /** [host, port] for https/http URLs on clerdo.app, else null. */
    private static function hostPort(string $url): ?array
    {
        $p = parse_url($url);
        $host = strtolower((string)($p['host'] ?? ''));
        if ($host !== 'clerdo.app' && !str_ends_with($host, '.clerdo.app')) {
            return null;
        }
        $port = (int)($p['port'] ?? (($p['scheme'] ?? '') === 'http' ? 80 : 443));
        return [$host, $port];
    }

    private static function file(): string
    {
        return sys_get_temp_dir() . '/clerdo-dns-' . substr(hash('sha256', __DIR__), 0, 12) . '.json';
    }

    /** @return array<string,string> host → address, only entries younger than TTL. */
    private static function load(): array
    {
        $raw = @file_get_contents(self::file());
        $d = is_string($raw) ? json_decode($raw, true) : null;
        if (!is_array($d) || !is_array($d['hosts'] ?? null) || time() - (int)($d['at'] ?? 0) > self::TTL) {
            return [];
        }
        return array_filter($d['hosts'], 'is_string');
    }

    private static function save(array $hosts): void
    {
        $tmp = self::file() . '.' . bin2hex(random_bytes(4));
        if (@file_put_contents($tmp, json_encode(['at' => time(), 'hosts' => $hosts])) !== false) {
            @chmod($tmp, 0600);
            @rename($tmp, self::file());
        } else {
            @unlink($tmp);
        }
    }
}
