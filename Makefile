TEX=paper/main.tex

pdf:
	latexmk -pdf $(TEX)

clean:
	latexmk -C
