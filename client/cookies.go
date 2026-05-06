package client

import (
	"encoding/json"
	"net/http"
	"net/http/cookiejar"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"
)

const persistentCookieJarVersion = 1

type persistentCookieJarData struct {
	Version int                              `json:"v"`
	Cookies map[string]persistentCookieEntry `json:"cookies"`
}

type persistentCookieEntry struct {
	URL    string           `json:"url"`
	Cookie persistentCookie `json:"cookie"`
}

type persistentCookie struct {
	Name     string     `json:"name"`
	Value    string     `json:"value"`
	Path     string     `json:"path,omitempty"`
	Domain   string     `json:"domain,omitempty"`
	Expires  *time.Time `json:"expires,omitempty"`
	Secure   bool       `json:"secure,omitempty"`
	HttpOnly bool       `json:"http_only,omitempty"`
	SameSite int        `json:"same_site,omitempty"`
}

// PersistentCookieJar persists HTTP cookies to disk for API affinity.
type PersistentCookieJar struct {
	mu   sync.Mutex
	path string
	jar  *cookiejar.Jar
	data persistentCookieJarData
	now  func() time.Time
}

func NewPersistentCookieJar(path string) (*PersistentCookieJar, error) {
	jar, err := cookiejar.New(nil)
	if err != nil {
		return nil, err
	}

	pj := &PersistentCookieJar{
		path: path,
		jar:  jar,
		data: persistentCookieJarData{
			Version: persistentCookieJarVersion,
			Cookies: make(map[string]persistentCookieEntry),
		},
		now: time.Now,
	}
	pj.load()
	return pj, nil
}

func (pj *PersistentCookieJar) SetCookies(u *url.URL, cookies []*http.Cookie) {
	pj.mu.Lock()
	defer pj.mu.Unlock()

	pj.jar.SetCookies(u, cookies)
	changed := false
	now := pj.now()
	for _, cookie := range cookies {
		if cookie == nil || cookie.Name == "" {
			continue
		}
		key := persistentCookieKey(u, cookie)
		if cookie.MaxAge < 0 || cookieExpired(cookie, now) {
			delete(pj.data.Cookies, key)
			changed = true
			continue
		}
		pj.data.Cookies[key] = persistentCookieEntry{
			URL:    persistentCookieURL(u),
			Cookie: cookieToPersistent(cookie, now),
		}
		changed = true
	}
	if changed {
		pj.save()
	}
}

func (pj *PersistentCookieJar) Cookies(u *url.URL) []*http.Cookie {
	pj.mu.Lock()
	defer pj.mu.Unlock()
	return pj.jar.Cookies(u)
}

func (pj *PersistentCookieJar) load() {
	if pj.path == "" {
		return
	}

	raw, err := os.ReadFile(pj.path)
	if err != nil {
		return
	}

	var data persistentCookieJarData
	if err := json.Unmarshal(raw, &data); err != nil || data.Version != persistentCookieJarVersion || data.Cookies == nil {
		return
	}

	now := pj.now()
	for key, entry := range data.Cookies {
		u, err := url.Parse(entry.URL)
		if err != nil {
			continue
		}
		cookie := persistentToCookie(entry.Cookie)
		if cookie.Name == "" || cookieExpired(cookie, now) {
			continue
		}
		pj.jar.SetCookies(u, []*http.Cookie{cookie})
		pj.data.Cookies[key] = entry
	}
}

func (pj *PersistentCookieJar) save() {
	if pj.path == "" {
		return
	}
	if err := os.MkdirAll(filepath.Dir(pj.path), 0o700); err != nil {
		return
	}

	raw, err := json.MarshalIndent(pj.data, "", "  ")
	if err != nil {
		return
	}
	raw = append(raw, '\n')

	tmp := pj.path + ".tmp"
	if err := os.WriteFile(tmp, raw, 0o600); err != nil {
		return
	}
	if err := os.Chmod(tmp, 0o600); err != nil {
		_ = os.Remove(tmp)
		return
	}
	if err := os.Remove(pj.path); err != nil && !os.IsNotExist(err) {
		_ = os.Remove(tmp)
		return
	}
	if err := os.Rename(tmp, pj.path); err != nil {
		_ = os.Remove(tmp)
		return
	}
	_ = os.Chmod(pj.path, 0o600)
}

func cookieToPersistent(cookie *http.Cookie, now time.Time) persistentCookie {
	var expires *time.Time
	if !cookie.Expires.IsZero() {
		t := cookie.Expires.UTC()
		expires = &t
	} else if cookie.MaxAge > 0 {
		t := now.Add(time.Duration(cookie.MaxAge) * time.Second).UTC()
		expires = &t
	}

	return persistentCookie{
		Name:     cookie.Name,
		Value:    cookie.Value,
		Path:     cookie.Path,
		Domain:   cookie.Domain,
		Expires:  expires,
		Secure:   cookie.Secure,
		HttpOnly: cookie.HttpOnly,
		SameSite: int(cookie.SameSite),
	}
}

func persistentToCookie(cookie persistentCookie) *http.Cookie {
	result := &http.Cookie{
		Name:     cookie.Name,
		Value:    cookie.Value,
		Path:     cookie.Path,
		Domain:   cookie.Domain,
		Secure:   cookie.Secure,
		HttpOnly: cookie.HttpOnly,
		SameSite: http.SameSite(cookie.SameSite),
	}
	if cookie.Expires != nil {
		result.Expires = cookie.Expires.UTC()
	}
	return result
}

func cookieExpired(cookie *http.Cookie, now time.Time) bool {
	return !cookie.Expires.IsZero() && !cookie.Expires.After(now)
}

func persistentCookieKey(u *url.URL, cookie *http.Cookie) string {
	host := strings.ToLower(u.Hostname())
	domain := strings.ToLower(cookie.Domain)
	if domain == "" {
		domain = host
	}
	path := cookie.Path
	if path == "" {
		path = defaultCookiePath(u.EscapedPath())
	}
	return strings.Join([]string{
		strings.ToLower(u.Scheme),
		host,
		domain,
		path,
		cookie.Name,
	}, "\t")
}

func persistentCookieURL(u *url.URL) string {
	copy := *u
	copy.RawQuery = ""
	copy.Fragment = ""
	if copy.Path == "" {
		copy.Path = "/"
	}
	return copy.String()
}

func defaultCookiePath(path string) string {
	if path == "" || path[0] != '/' {
		return "/"
	}
	lastSlash := strings.LastIndex(path, "/")
	if lastSlash <= 0 {
		return "/"
	}
	return path[:lastSlash]
}
