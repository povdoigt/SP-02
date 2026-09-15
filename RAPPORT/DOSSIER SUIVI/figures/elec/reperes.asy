// ---------------------------------------------------------------
//  Reperes de la fusee SP-02 : repere Terre, repere fusee, repere capteur
//  Compilation :  asy -f pdf reperes.asy
// ---------------------------------------------------------------

import three;
import solids;

// settings.outformat = "pdf";
settings.prc      = false;   // pas de 3D interactive dans le PDF
settings.render   = 8;       // 0 = sortie vectorielle ; passer a 8 si besoin

size(13cm);
currentprojection = orthographic(5, -4, 3);
currentlight      = light(gray(0.55), specular=black, (0.5, -1, 1));

// ----------------------- parametres --------------------------------
real elevation = 45;                    // angle de l'axe fusee / horizontale
real azimut    = 60;                    // orientation dans le plan horizontal

real R  = 0.15;                         // rayon du corps
real L  = 2.00;                         // longueur du corps
real Hc = 0.65;                         // hauteur de la coiffe

triple axe = (cos(radians(elevation))*cos(radians(azimut)),
              cos(radians(elevation))*sin(radians(azimut)),
              sin(radians(elevation)));

triple cap = (cos(radians(azimut/2)), sin(radians(azimut/2)), 0);

triple culot = O + 1.5*axe;         // position du culot
triple haut  = culot + L*axe;           // jonction corps / coiffe

// base orthonormee liee a la fusee
triple xF = unit(cross(axe, Z));
triple yF = cross(axe, xF);

// ----------------------- plan du sol -------------------------------
path3 sol = (-0.8,-0.8,0) -- (3.0,-0.8,0) -- (3.0,3.0,0) -- (-0.8,3.0,0) -- cycle;
//draw(surface(sol), gray(0.95) + opacity(0.5));
draw(sol, gray(0.7)+linewidth(0.3pt));

// ----------------------- repere Terre ------------------------------
pen pT = gray(0.25) + linewidth(0.8pt);
draw(O -- 1.3X, pT, Arrow3(size=4mm));
draw(O -- 1.3Y, pT, Arrow3(size=4mm));
draw(O -- 1.3Z, pT, Arrow3(size=4mm));
label("$x_T$", 1.4X, SW, gray(0.25));
label("$y_T$", 1.2Y, NW, gray(0.25));
label("$z_T$", 1.4Z,  N, gray(0.25));
label("$\mathcal{R}_T$", O, SW, gray(0.25));

pen pA = gray(0.7) + linewidth(0.4pt);
triple A = O + 4.5*axe;
triple Proj_A = A - (dot(A - O, Z) / dot(Z, Z)) * Z;
draw(O -- A,  pA, Arrow3(size=2mm));
draw(A -- Proj_A, dotted);
draw(O -- Proj_A, dotted);
draw(O -- cap*sqrt(dot(Proj_A, Proj_A))*2/3, dotted);
draw(O -- 1.20*cap, orange, Arrow3(size=2mm));
label("$cap_0$", 1.2*cap, S, orange);

// ----------------------- corps de la fusee -------------------------
// solides de revolution : silhouettes et parties cachees exactes
revolution corps  = cylinder(culot, R, L,  axe);
revolution coiffe = cone    (haut,  R, Hc, axe);

draw(surface(corps),  gray(0.82) + opacity(0.5));
draw(surface(coiffe), gray(0.72) + opacity(0.5));

// ailerons : quadrilateres plans contenant l'axe
void aileron(triple d) {
  path3 p = (culot + R*d)
         -- (culot + 0.45*d - 0.1*axe)
         -- (culot + 0.45*d + 0.15*axe)
         -- (culot + R*d + 0.55*axe) -- cycle;
  draw(surface(p), gray(0.62));
  draw(p, gray(0.35)+linewidth(0.3pt));
}
aileron(xF); aileron(-xF); aileron(yF); aileron(-yF);

// ----------------------- repere fusee ------------------------------
triple oF = culot + 0.4*L*axe;
pen pF = blue + linewidth(0.8pt);
draw(oF -- oF + 0.8*unit(xF + yF),  pF, Arrow3(size=4mm));
draw(oF -- oF + 0.8*unit(yF - xF),  pF, Arrow3(size=4mm));
draw(oF -- oF + 0.8*axe, pF, Arrow3(size=4mm));
label("$z_F$", oF + 0.8*unit(xF + yF), SE, blue);
label("$x_F$", oF + 0.8*unit(yF - xF),  E, blue);
label("$y_F$", oF + 0.8*axe + 0.4*yF , NW, blue);
// label("$x_F$", oF + 0.9*xF - 0.2*axe - 0.15*yF, SE, blue);
// label("$y_F$", oF + 0.8*yF                    ,  E, blue);
// label("$z_F$", oF + 0.8*axe                   , SE, blue);
label("$\mathcal{R}_F$", oF + 0.5*axe - 0.3*yF, SW, blue);

// ----------------------- repere capteur ----------------------------
// centrale inertielle : meme axe longitudinal, rotation de 90 deg autour de z
// triple oC = culot + 0.80*L*axe;
// triple xC = yF, yC = -xF;
// pen pC = orange + linewidth(0.8pt);
// draw(oC -- oC + 0.60*xC,  pC, Arrow3(size=3.5mm));
// draw(oC -- oC + 0.60*yC,  pC, Arrow3(size=3.5mm));
// draw(oC -- oC + 0.60*axe, pC, Arrow3(size=3.5mm));
// label("$x_C$", oC + 0.68*xC,  SE, orange);
// label("$y_C$", oC + 0.68*yC,  W,  orange);
// label("$z_C$", oC + 0.68*axe, SE, orange);
// label("$\mathcal{R}_C$", oC, SW, orange);

// ----------------------- angle d'elevation -------------------------
// arc entre la verticale et l'axe longitudinal
draw(arc(O, 1.20*unit(Proj_A), 1.20*axe), red, Arrow3(size=3mm));
label("$\theta$", 1.20*unit(unit(Proj_A) + axe), NE, red);

draw(arc(O, 0.9*X, 0.9*cap), orange, Arrow3(size=3mm));
label("$\phi$", 0.8*unit(X + 2*cap), W, orange);

draw(arc(O, 1.20*cap, 1.20*unit(Proj_A)), orange, Arrow3(size=3mm));
label("$\Delta\phi$", 1.20*unit(cap + unit(Proj_A)), W, orange);
