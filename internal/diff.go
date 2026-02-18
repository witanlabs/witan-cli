package internal

import (
	"fmt"
	"image"
	"image/color"
)

const (
	// innerRadius is the distance for the inner (black) outline stroke.
	innerRadius = 1
	// outerRadius is the distance for the outer (white) outline stroke.
	outerRadius = 2
)

var (
	strokeInner = color.RGBA{R: 0, G: 0, B: 0, A: 255}       // black
	strokeOuter = color.RGBA{R: 255, G: 255, B: 255, A: 255} // white
)

// DiffImages compares two images pixel-by-pixel and returns a diff image.
// Changed pixels show the "after" value at full color, surrounded by a
// black+white double-stroke outline. Unchanged pixels are desaturated and
// dimmed. Returns the count of changed pixels.
func DiffImages(before, after image.Image) (*image.RGBA, int, error) {
	if before.Bounds() != after.Bounds() {
		bb := before.Bounds()
		ab := after.Bounds()
		return nil, 0, fmt.Errorf(
			"image dimensions differ: before is %d×%d, after is %d×%d — use the same --range and --dpr for both renders",
			bb.Dx(), bb.Dy(), ab.Dx(), ab.Dy(),
		)
	}

	bounds := after.Bounds()
	w := bounds.Dx()
	h := bounds.Dy()

	// Pass 1: build changed-pixel mask
	mask := make([]bool, w*h)
	changed := 0
	for y := bounds.Min.Y; y < bounds.Max.Y; y++ {
		for x := bounds.Min.X; x < bounds.Max.X; x++ {
			br, bg, bb, ba := before.At(x, y).RGBA()
			ar, ag, ab, aa := after.At(x, y).RGBA()
			if br != ar || bg != ag || bb != ab || ba != aa {
				mask[(y-bounds.Min.Y)*w+(x-bounds.Min.X)] = true
				changed++
			}
		}
	}

	// Pass 2: for each unchanged pixel, compute squared distance to nearest changed pixel.
	// We only need to distinguish: inner stroke (<=innerRadius), outer stroke (<=outerRadius), or neither.
	// Use value 0=changed, 1=inner, 2=outer, 3=none.
	const (
		zChanged = 0
		zInner   = 1
		zOuter   = 2
		zNone    = 3
	)
	zone := make([]uint8, w*h)
	for i := range zone {
		if mask[i] {
			zone[i] = zChanged
		} else {
			zone[i] = zNone
		}
	}

	if changed > 0 {
		r := outerRadius
		ir2 := innerRadius * innerRadius
		or2 := outerRadius * outerRadius
		for y := 0; y < h; y++ {
			for x := 0; x < w; x++ {
				idx := y*w + x
				if zone[idx] == zChanged {
					continue
				}
				yMin := max(0, y-r)
				yMax := min(h-1, y+r)
				xMin := max(0, x-r)
				xMax := min(w-1, x+r)
				minDist2 := or2 + 1 // sentinel
				for ny := yMin; ny <= yMax; ny++ {
					for nx := xMin; nx <= xMax; nx++ {
						if mask[ny*w+nx] {
							dx := nx - x
							dy := ny - y
							d2 := dx*dx + dy*dy
							if d2 < minDist2 {
								minDist2 = d2
							}
						}
					}
				}
				if minDist2 <= ir2 {
					zone[idx] = zInner
				} else if minDist2 <= or2 {
					zone[idx] = zOuter
				}
			}
		}
	}

	// Pass 3: render
	result := image.NewRGBA(bounds)
	for y := bounds.Min.Y; y < bounds.Max.Y; y++ {
		for x := bounds.Min.X; x < bounds.Max.X; x++ {
			idx := (y-bounds.Min.Y)*w + (x - bounds.Min.X)
			ar, ag, ab, aa := after.At(x, y).RGBA()

			switch zone[idx] {
			case zChanged:
				result.SetRGBA(x, y, color.RGBA{
					R: uint8(ar >> 8),
					G: uint8(ag >> 8),
					B: uint8(ab >> 8),
					A: uint8(aa >> 8),
				})
			case zInner:
				result.SetRGBA(x, y, strokeInner)
			case zOuter:
				result.SetRGBA(x, y, strokeOuter)
			default:
				// Unchanged: desaturate + dim
				r8 := float64(ar >> 8)
				g8 := float64(ag >> 8)
				b8 := float64(ab >> 8)
				gray := 0.299*r8 + 0.587*g8 + 0.114*b8
				dimmed := 0.3*gray + 0.7*255
				d := uint8(dimmed)
				result.SetRGBA(x, y, color.RGBA{R: d, G: d, B: d, A: uint8(aa >> 8)})
			}
		}
	}

	return result, changed, nil
}

// FormatDiffSummary returns a human-readable diff summary string.
func FormatDiffSummary(changed, total int) string {
	if changed == 0 {
		return "diff: no changes"
	}
	pct := float64(changed) / float64(total) * 100
	if pct < 0.1 {
		return fmt.Sprintf("diff: %d pixels changed (<0.1%%)", changed)
	}
	return fmt.Sprintf("diff: %d pixels changed (%.1f%%)", changed, pct)
}
