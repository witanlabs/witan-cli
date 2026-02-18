package internal

import (
	"image"
	"image/color"
	"strings"
	"testing"
)

func solidImage(w, h int, c color.RGBA) *image.RGBA {
	img := image.NewRGBA(image.Rect(0, 0, w, h))
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			img.SetRGBA(x, y, c)
		}
	}
	return img
}

func TestDiffImages_Identical(t *testing.T) {
	c := color.RGBA{R: 100, G: 150, B: 200, A: 255}
	img := solidImage(4, 4, c)

	result, changed, err := DiffImages(img, img)
	if err != nil {
		t.Fatal(err)
	}
	if changed != 0 {
		t.Errorf("expected 0 changed, got %d", changed)
	}

	// All pixels should be dimmed (grayish, not original color)
	px := result.RGBAAt(0, 0)
	if px.R != px.G || px.G != px.B {
		t.Errorf("expected grayscale pixel, got R=%d G=%d B=%d", px.R, px.G, px.B)
	}
	// Should be brighter than original gray due to 70% white blend
	if px.R < 200 {
		t.Errorf("expected dimmed pixel > 200 (bright gray), got %d", px.R)
	}
}

func TestDiffImages_FullyDifferent(t *testing.T) {
	before := solidImage(4, 4, color.RGBA{R: 0, G: 0, B: 0, A: 255})
	after := solidImage(4, 4, color.RGBA{R: 255, G: 0, B: 0, A: 255})

	result, changed, err := DiffImages(before, after)
	if err != nil {
		t.Fatal(err)
	}
	if changed != 16 {
		t.Errorf("expected 16 changed, got %d", changed)
	}

	// All pixels should be the after color (red)
	px := result.RGBAAt(0, 0)
	if px.R != 255 || px.G != 0 || px.B != 0 {
		t.Errorf("expected red pixel, got R=%d G=%d B=%d", px.R, px.G, px.B)
	}
}

func TestDiffImages_SinglePixelChange(t *testing.T) {
	c := color.RGBA{R: 128, G: 128, B: 128, A: 255}
	// Use 20x20 so there are pixels far from the change (well beyond outline radius)
	before := solidImage(20, 20, c)
	after := solidImage(20, 20, c)
	after.SetRGBA(10, 10, color.RGBA{R: 255, G: 0, B: 0, A: 255})

	result, changed, err := DiffImages(before, after)
	if err != nil {
		t.Fatal(err)
	}
	if changed != 1 {
		t.Errorf("expected 1 changed, got %d", changed)
	}

	// The changed pixel should be full red
	px := result.RGBAAt(10, 10)
	if px.R != 255 || px.G != 0 || px.B != 0 {
		t.Errorf("changed pixel: expected red, got R=%d G=%d B=%d", px.R, px.G, px.B)
	}

	// A pixel 1px away should be inner stroke (black)
	inner := result.RGBAAt(10, 9) // 1px away, within innerRadius=1
	if inner.R != 0 || inner.G != 0 || inner.B != 0 {
		t.Errorf("inner stroke pixel: expected black, got R=%d G=%d B=%d", inner.R, inner.G, inner.B)
	}

	// A pixel 2px away (diagonal) should be outer stroke (white)
	outer := result.RGBAAt(9, 9) // sqrt(2) ≈ 1.4px away, between innerRadius=1 and outerRadius=2
	if outer.R != 255 || outer.G != 255 || outer.B != 255 {
		t.Errorf("outer stroke pixel: expected white, got R=%d G=%d B=%d", outer.R, outer.G, outer.B)
	}

	// A far pixel should be dimmed gray
	ux := result.RGBAAt(0, 0) // 14px away, well beyond radius
	if ux.R != ux.G || ux.G != ux.B {
		t.Errorf("unchanged pixel: expected grayscale, got R=%d G=%d B=%d", ux.R, ux.G, ux.B)
	}
}

func TestDiffImages_DimensionMismatch(t *testing.T) {
	before := solidImage(4, 4, color.RGBA{A: 255})
	after := solidImage(5, 3, color.RGBA{A: 255})

	_, _, err := DiffImages(before, after)
	if err == nil {
		t.Fatal("expected error for dimension mismatch")
	}
	// Should mention dimensions
	if got := err.Error(); !strings.Contains(got, "4×4") || !strings.Contains(got, "5×3") {
		t.Errorf("error should mention both dimensions, got: %s", got)
	}
}

func TestFormatDiffSummary(t *testing.T) {
	tests := []struct {
		changed int
		total   int
		want    string
	}{
		{0, 100, "diff: no changes"},
		{42, 14000, "diff: 42 pixels changed (0.3%)"},
		{1, 1000000, "diff: 1 pixels changed (<0.1%)"},
		{500, 1000, "diff: 500 pixels changed (50.0%)"},
	}
	for _, tt := range tests {
		got := FormatDiffSummary(tt.changed, tt.total)
		if got != tt.want {
			t.Errorf("FormatDiffSummary(%d, %d) = %q, want %q", tt.changed, tt.total, got, tt.want)
		}
	}
}
