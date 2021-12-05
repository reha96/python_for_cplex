/*********************************************
 * OPL 20.1.0.0 Model
 * Author: Reha
 * Creation Date: Feb 15, 2021 at 9:02:16 PM
 *********************************************/
/* This code can be used to compute the number of types: each color c resembles a different type */
int NR_PublicGoods=...;
int NR_PrivateGoods=...;
int NR_Goods=NR_PublicGoods+NR_PrivateGoods;

int NR_Observations = ...;
range households = 1..NR_Observations;
range Goods=1..NR_Goods;

float AllP[households][1..NR_Goods]=...;
float AllQ[households][1..NR_Goods]=...;
float Income[households]=...;
range colors = 1..22; //let's assume that there are XX types at max

dvar boolean y[colors]; // indicator to tell if the type is empty or not
dvar boolean s[households,colors]; // indicator to tell if household (i) is in type (c) or not 
dexpr int NR_Types=sum(c in colors) y[c];

minimize NR_Types;

subject to {

// Constraints for proper coloring
forall(i in households) sum(c in colors) s[i,c]  == 1; //each individual belongs to exactly one partition (p.13, step 2)
forall(i in households, c in colors) s[i,c] <= y[c];
forall(i in households, j in households, c in colors: i !=j) ((sum(q in Goods) (AllP[i][q]*AllQ[i][q]) >= sum(q in Goods) (AllP[i][q]*AllQ[j][q])) 
&& (sum(q in Goods) (AllP[j][q]*AllQ[j][q]) >= sum(q in Goods) (AllP[j][q]*AllQ[i][q]))) => s[i,c] + s[j,c] <= 1;
}
