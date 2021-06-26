%------------------------------------------------------------------------
% A Software Code for the Estimation of Thermodynamic and Rate           %
% Parameters of Surface Reactions                                        %
% Author: Mahmoud Moqadam, M.Sc of Amirkabir University of Technology    %
% Contact information: Moqadam@aut.ac.ir (Mahmoud Moqadam)               %
% Date: 2013                                                             %
%% ------------------------------------------------------------------------
% Application of the Code:
% Calculation of
%   1. Thermodynamic parameters
%   2. Rate parameters
%   3. Flow rate 
%   4. Surface coverage
% Two calculation modules in the Code:
% Module 1: Calculation of thermodynamic parameters
% Module 2: Calculation of rate parameters
% Module 3: Calculation of kinetic reaction Equation
% The variables are defined here:
%
% NUMBER_REACTIONS    Number of surface reactions in the mechanism
% A_forward reaction  A-factor of forward reaction
% A_reverse reaction  A-factor of reverse reaction
% E_forward           Activation energy of forward reaction for each reaction type
% E_reverse           Activation energy of forward reaction for each reaction type
% coordination_number Coordination number of the adsorbate bound to catalyst
% (On-top= 1, bridge=2)
% binding_strength Strength of binding (weak=l, medium=2, strong=3)
% Medium bounded adsorbate AB, Is with the A end down:1 or coordinated via
% both A and B:2 or Symmetric:3
% coordination_num    Coordination number of the adsorbate bound to catalyst
% coordination_num_X  Coordination number of the X atom bound to A
% coordination_num_Y  Coordination number of the Y atom bound to B
% Q_species           Heats of adsorption
% Q_metal_atom        atomic heat of chemisorption
% Q_metal_atom_0      first estimation for atomic heat of chemisorption
% Q0_metal_atom_A     Uni-coordinated atomic heat of chemisorption for A
% Q0_metal_atom_B     Uni-coordinated atomic heat of chemisorption for B
% Q0_metal_atom_X     Uni-coordinated atomic heat of chemisorption for X
% Q0_metal_atom_Y     Uni-coordinated atomic heat of chemisorption for Y
% Q_species_total     Molecular chemisorption enthalpy for intermediate binding
% Q_metal_atom_A      Metal-atom bonding for A
% Q_metal_atom_B      Metal-atom bonding for B
% Q_metal_atom_X      Metal-atom bonding for X
% Q_metal_atom_Y      Metal-atom bonding for Y
% D_AB                Gas phase dissociation energy between atom or group A and B
% D_AB_red            Reduced Gas phase dissociation energy between atom or group A and B 
% Raction_type        Type of each reaction
% component_num       Number of componenet in reaction
% main_atom_num       Number of involved atoms
% bonding_structure   structure of bond: mono, di, poly 
% R                   gasous constant
% v                   Volume value
% p                   pressure value
% flo                 input folw rate value
% k0                  pre-exponential value
% k                   Rate constant
% xexp                Experiment data
% bsth                Binding strength in excel
% bsre                bonding structure in excel
% D_Q                 Gas phase dissociation energy in excel
% Dr_Q                Reduce Gas phase dissociation energy in excel
% rtype               reaction type  
% rdisc               reaction discription  
% D_E                 D_AB for Activation energy part in excel
% T                   temprature of data set
% Q0_A                Matrix for Q0_metal_atom 
% Q_calculation       matrix Q_meta_atom in excel for matrix component
% Q_calculation_E     matrix Q_meta_atom in excel for matrix reaction
% No_E                number for each reaction in excel
% E_AB                Activation energy, forward and reverse
% dH                  Enthalpy
% x0                  first estimation for TETAs and Flow Rates
% x                   TETAs and Flow Rates  
% xf                  flow rate (output from solver)
% -----------------------------------------------------------------------
%% Start the Program
function Final_Micro_Kinetic_Q_and_Reac_finished
clear
clc
disp('*** diffrent kind of binding strength ***');
disp('@ Weak: CO, H2, H2O, CH4, CH3OH, CH3CHO, CH2CO, CH3CH2OH, NO, O2');
disp('@ Strong: OH, CH, CH2, CH3O, CCH3, CHCH3, CHCHO, CH3CH2O');
disp('@ Medium: CH3, HCO, CH3CO, CH2CHO, CH2CH3');
disp('@ Weakly bounded two atoms to surface: CHCH, CH2CH2, CH3CH3, CCH, CCH2, CHCH2, CHCO, CH2CO');
disp('@ specific bond like A-X-B');

format long
% gasous constant
R = 8.314;

% read The matrix of Volume value
v=xlsread('testdata1.xlsx',3,'A8');

% read The matrix of pressure value
p=xlsread('testdata1.xlsx',1,'D2:D5');

% read The matrix of input folw rate value
flo=xlsread('testdata1.xlsx',1,'F2:F4');

% read The matrix of pre-exponential value
k0=xlsread('testdata1.xlsx',1,'A2:B5');

% read The matrix of Experiment data
xexp=xlsread('testdata1.xlsx',1,'B18:B24');

% read The matrix of Binding strength
bsth=xlsread('testdata1.xlsx',3,'B11:B13');

% read The matrix of bonding structure
bsre = xlsread('testdata1.xlsx',3,'C11:C13');

% read The matrix of D_AB for Surface Species
D_Q = xlsread('testdata1.xlsx',3,'D11:D13');

% read The matrix of D_AB Reduce for Surface Species
Dr_Q = xlsread('testdata1.xlsx',3,'E11:E13');

% read The matrix of D_AXB for A-X-B Surface Species
D_AXB = xlsread('testdata1.xlsx',3,'P11:Q13');

% read the matrix of reaction type
rtype = xlsread('testdata1.xlsx',3,'B31:B34');

% read the matrix of reaction discription
rdisc = xlsread('testdata1.xlsx',3,'C31:C34');

% read The matrix of D_AB for Activation energy part
D_E = xlsread('testdata1.xlsx',3,'D31:D34');

% Heat of chemisorption Q_X, Y, Z, F for Activation energy Propagation part
Q_XYZF = xlsread('testdata1.xlsx',3,'F31:I34');

% read The matrix of D_X, Y, Z, F for Activation energy Propagation part
D_XYZF = xlsread('testdata1.xlsx',3,'J31:M34');

%first estimation for TETAs and Flow Rates
x0=[1.0E-10;2.0E-7;1.0E-1;8.9E-1;8;2;6];

% read temprature of data set
disp('Value of Temprature');
T = input('Temprature : ');
disp('-----------------');
disp('Surface Species');

% read number of main atom involved (H, C, O)
main_atom_num = xlsread('testdata1.xlsx',3,'E2');
display(main_atom_num);
%Q0_A = [1,main_atom_num];
disp('-----------------');

% read number of components
component_num=xlsread('testdata1.xlsx',3,'A14');

% number of reaction involved
main_reaction_num = xlsread('testdata1.xlsx',3,'A37');

% read first estimations for Heat of chemisorption and Coordination numbers
disp('*** first input value of H then C and O and continue ... ***');
Q_metal_atom_0 = xlsread('testdata1.xlsx',3,'B3:D3');
display(Q_metal_atom_0)

% read Coordination numbers for component
%coordination_num = xlsread('testdata1.xlsx',3,'B4:D4');
coordination_num = xlsread('testdata1.xlsx',3,'H11:I13');
display(coordination_num)
disp('-----------------');

% read Heat of chemisorption (Q_metal_atom) and build Q_calc matrix for components
Q_calculation = xlsread('testdata1.xlsx',3,'F11:G13');
Q_calc=Q_calculation;
display(Q_calc)

% read m , n and Heat of chemisorption (Q_metal_atom) and build Q_calc matrix for components AXm-BYn
Q_calc_XY = xlsread('testdata1.xlsx',3,'J11:O13');
display(Q_calc_XY)

% read a number of component for every reaction to determine Q_AB
No_E = xlsread('testdata1.xlsx',3,'N31:N34');
display(No_E)

% read a number of component for every reaction to determine Q_A
No_A = xlsread('testdata1.xlsx',3,'F31:F34');

% read a number of component for every reaction to determine Q_B
No_B = xlsread('testdata1.xlsx',3,'G31:G34');

% build Q_A (Q_metal_atom) matrix by first estimation
Q_A=Q_metal_atom_0;
display(Q_A)

% rate constant for reactions
k=zeros(main_reaction_num,2);

% Activation energy each species in Reaction part
E_AB(main_reaction_num,2)=zeros;

% the matrix in which place Q_atom or component for reactions
Q_calculation_E(main_reaction_num,2)=zeros;
Q_calc_E=Q_calculation_E;

r=1;
G(r,7)=zeros;
Q(r,3)=zeros;
QQQ(r,3)=zeros;
%%
% optimizer procedure
options=optimset('MaxFunEvals',10000);
Q_metal_atom=fminsearch(@obj,Q_metal_atom_0,options);
disp('-***-');
display(Q_metal_atom)
disp('-***-');

% function of error to optimize        
    function error=obj(Q_metal_atom)
        
% Calculation of Metal–atom bonding Q0_A
disp('-----------------');
disp('-----------------');
display(Q_metal_atom);
disp('-----------------');
disp('-----------------');

% Determination of kind of Binding strength and Equations used to calculate heat of adsorptions
% Read all copmonent and determine kind of binding strength and bonding structure
Q_species = [1,component_num];
for j=1:component_num


    % replace new data of Q_metal_atom from optimizer in Q_calc and Q_calc_XY matrix for J th component
    display(j)
    display(Q_metal_atom)
    for iii=1:2
        for tt=1:main_atom_num
            if Q_calc(j,iii)==Q_A(tt)
                Q_calc(j,iii)=Q_metal_atom(tt);
            end
            if Q_calc_XY(j,iii)==Q_A(tt)
                Q_calc_XY(j,iii)=Q_metal_atom(tt);
            end            
        end
    end
      
    display(Q_calc)    
    display(Q_calc_XY)
    if j==3
     Q_A=Q_metal_atom;
    end
    display(Q_A)
    
    disp(['The value of J th component from Excel file is: ', num2str(j)]);
    disp('The Kind of Binding strength (weak =l, medium =2, strong =3, specific bond =4)');
    % Determination of J th binding strength
    binding_strength = bsth(j);
    display(binding_strength);

    if binding_strength == 1
        disp('The kind of bonding structure: A End Down : M-A-B=1 and two atoms (A and B) to the surface : AXm-BYn=2');
        % Determination of J th bonding structure
        bonding_structure = bsre(j);
        display(bonding_structure);
      
        if bonding_structure == 1
        disp('*** On-top Coordination of A and B With the A End Down. M-A-B ***');
        
        % Q0_metal_atom_A : copmare the value of J th component and determine which component is and give a value of Q0_A to it
        Q0_metal_atom_A=Q_calc(j,1)/(2-(1/coordination_num(j,1)));
     
        % give a D_AB to J th component        
        D_AB = D_Q(j);
        display(D_Q)
        display(D_AB)
        display(Q0_metal_atom_A)

        % calculation of Q_AB for J th component
        Q_species(j) = Q0_metal_atom_A^2/(Q0_metal_atom_A+D_AB);
        
        else
            disp('*** Diatomic molecules weakly bounded through two atoms (A and B) to the surface : AXm-BYn ***');
            Q_metal_atom_X = Q_calc_XY(j,1);
            Q_metal_atom_Y = Q_calc_XY(j,2);
            coordination_num_X = Q_calc_XY(j,5);
            Q0_metal_atom_X =Q_metal_atom_X/(2-(1/coordination_num_X));
            coordination_num_Y = Q_calc_XY(j,6);
            Q0_metal_atom_Y =Q_metal_atom_Y/(2-(1/coordination_num_Y));
            Q0_metal_atom_A = Q_calc(j,1)/(2-(1/coordination_num(j,1)));
            Q0_metal_atom_B = Q_calc(j,2)/(2-(1/coordination_num(j,2)));
            % give a D_AB Reduce to J th component                    
            D_AB_red = Dr_Q(j);
            m = Q_calc_XY(j,3);
            n = Q_calc_XY(j,4);
            a=(Q0_metal_atom_A)*(1-((Q0_metal_atom_X*m)/(Q0_metal_atom_A*m + Q0_metal_atom_X))^2);
            b=(Q0_metal_atom_B)*(1-((Q0_metal_atom_Y*n)/(Q0_metal_atom_B*n + Q0_metal_atom_Y))^2);
            
            % calculation of Q_AB for J th component
            Q_species(j)=(a*b*(a+b)+D_AB_red*(a-b)^2)/(a*b+D_AB_red*(a+b));
        end
    elseif binding_strength == 2
        disp('type of Medium bounded adsorbate AB, Is with the A end down:1 or coordinated via both A and B:2 or Symmetric:3 ');

        bonding_structure = bsre(j);
        
        if bonding_structure == 1
            disp('*** A Medium bounded adsorbate AB, with the A end down ***');
            
            % Q0_metal_atom_A: copmare the value of J th component and determine which component is and give a value of Q0_A to it
            Q0_metal_atom_A=Q_calc(j,1)/(2-(1/coordination_num(j,1)));
                    
            %Q_metal_atom_A : The value of QA-M for this species in medium type 
            Q_metal_atom_A=Q_calc(j,1);

            % give a D_AB Reduce to J th component
            D_AB_red = Dr_Q(j);
            display(Dr_Q)
            display(Dr_AB_red)
            
            % calculation of Q_AB for J th component       
            Q_species(j) = 0.5*(((Q0_metal_atom_A^2)/(Q0_metal_atom_A+D_AB_red))+((Q_metal_atom_A^2)/(Q_metal_atom_A+D_AB_red)));
            
        elseif bonding_structure == 2
            disp('*** Medium bounded adsorbates, where AB is coordinated via both A and B ***');
            
            %Q0_metal_atom_A: copmare the value of J th component and determine which component is and give a value of Q0_A to it
            Q0_metal_atom_A=Q_calc(j,1)/(2-(1/coordination_num(j,1)));

            %Q0_metal_atom_B: copmare the value of J th component and determine which component is and give a value of Q0_B to it
            Q0_metal_atom_B=Q_calc(j,2)/(2-(1/coordination_num(j,2)));
                                       
            display(Q0_metal_atom_A)
            display(Q0_metal_atom_B)
 
            % give a D_AB Reduce to J th component                    
            D_AB_red = Dr_Q(j);
            display(Dr_Q)
            display(Dr_AB_red)

            a=(Q0_metal_atom_A^2)*((Q0_metal_atom_A + 2*Q0_metal_atom_B)/(Q0_metal_atom_A + Q0_metal_atom_B)^2);
            b=(Q0_metal_atom_B^2)*((Q0_metal_atom_B + 2*Q0_metal_atom_A)/(Q0_metal_atom_A + Q0_metal_atom_B)^2);
            
            % calculation of Q_AB for J th component       
            Q_species(j) = (a*b*(a+b)+D_AB_red*(a-b)^2)/(a*b+D_AB_red*(a+b));
        else
            disp('*** Symmetric species ***');
            
            %Q0_metal_atom_A: copmare the value of J th component and determine which component is and give a value of Q0_A to it
            Q0_metal_atom_A=Q_calc(j,1)/(2-(1/coordination_num(j,1)));
                    
            % give a D_AB to J th component
            D_AB = D_Q(j);
            display(D_Q)
            display(D_AB)
            
            % calculation of Q_AB for J th component                
            Q_species(j) = 4.5*Q0_metal_atom_A^2/(3*Q0_metal_atom_A+8*D_AB);
        end
    elseif binding_strength == 3
        disp('*** A strong bounded adsorbate AB, with the A end down ***');
        
        %Q_metal_atom_A: the value of J th component from (Q_calc) in strong bond
        Q_metal_atom_A=Q_calc(j,1);
        display(Q_metal_atom_A)
        
        % give a D_AB to J th component        
        D_AB = D_Q(j);
        display(D_Q)
        display(D_AB)
        
        % calculation of Q_AB for J th component                
        Q_species(j)= Q_metal_atom_A^2/(Q_metal_atom_A+D_AB);
    else
        disp('*** A molecule where A and B are linked together via some atom X and a chelate structure is formed ***');
        
        %Q0_metal_atom_A: copmare the value of J th component and determine which component is and give a value of Q0_A to it
        Q0_metal_atom_A=Q_calc(j,1)/(2-(1/coordination_num(j,1)));

        %Q0_metal_atom_B: copmare the value of J th component and determine which component is and give a value of Q0_B to it
        Q0_metal_atom_B=Q_calc(j,2)/(2-(1/coordination_num(j,2)));

        % determine values of D_AXB and other constants
        D_AX = D_AXB(j,1);
        D_BX = D_AXB(j,2);
        a = Q0_metal_atom_A^2/(Q0_metal_atom_A+D_AX);
        b = Q0_metal_atom_B^2/(Q0_metal_atom_B+D_BX);
        a_X = (a^2)*(a+2*b)/(a+b)^2;
        b_X = (b^2)*(b+2*a)/(b+a)^2;
        
        % calculation of Q_AB for J th component                
        Q_species(j)= a_X + b_X;
    end
    disp('-----------------');
end
% Show the values of Q_AB of componenets
display(Q_species);
disp('-----------------');

%%
disp('*** Activation Energy ***');  
disp('AB* is the Adsorbate Componenet');
disp('AB(g) is the Gaseous Copmonenet');

% Determination of kind of reaction type and discription used to calculate Activation Energy
% Read all reaction and determine type and discription
for i=1:main_reaction_num
    disp('type of reactions : ');
    disp('Gaseous reactions         AB(g)               = 1');
    disp('Surface reactions         AB*                 = 2');
    disp('One component reactions   AB                  = 3');
    disp('Disproportation reactions X* + Y* --> Z* + F* = 4');
display(i)
    % Determination of i th reaction type
    reaction_type = rtype(i);
    display(rtype);
    display(reaction_type);    
    disp('-----------------');

    % give a D_AB to i th component
    D_AB = D_E(i);

    % Q_metal_atom_A  and Q_metal_atom_B : The value of QA-M for this species
    % 1- Q_metal_atom_A  and Q_metal_atom_B are component so choose its heat of adsorbtion
    for t=1:component_num
        if No_A(i)==t
            Q_calculation_E(i,1)=Q_species(t);
        end
        if No_B(i)==t
            Q_calculation_E(i,2)=Q_species(t);
        end
    end

    % 2- Q_metal_atom_A  and Q_metal_atom_B are Atom so choose its heat of adsorbtion
    % 100 is just a symbol for counter (Example: for No. H=101=Q(1), C=102=Q(2), O=103=Q(3))
    for h=1:main_atom_num
        if No_A(i)==(h+100)
            Q_calculation_E(i,1)=Q_metal_atom(h);
        end
        if No_B(i)==(h+100)
            Q_calculation_E(i,2)=Q_metal_atom(h);
        end
    end
    display(Q_calculation_E)
    Q_calc_E=Q_calculation_E;
    Q_metal_atom_A=Q_calc_E(i,1);
    Q_metal_atom_B=Q_calc_E(i,2);
    display(Q_calc_E)
    display(Q_metal_atom_A)
    display(Q_metal_atom_B)
    
    % Q_AB : The value of Q-AB for this species (i th)
    % The No. of each reaction show the J the main component number (AB)
    % Q_AB = 0 when there are not any chemisorption
    for ttt=1:component_num
        if No_E(i)== ttt
           Q_AB=Q_species(ttt);
           break
        else
           Q_AB = 0;
        end
    end

    disp('Q_AB is the value of each component');
    display(Q_AB)
    disp('-----------------');
    
    % Enthaply of i th reaction
    dH = D_AB + Q_AB - Q_metal_atom_A - Q_metal_atom_B;
    
    if reaction_type==1
        disp('Gaseous reactions   AB(g)  ');
        disp('Choose reactions disciption : ');
        disp('Disscosiation to Gas-phase reaction  : AB(g) --> A* + B* = 1');
        disp('Recombination to Gas-phase reactions : A* + B* --> AB(g) = 2');

        % Determination of i th reaction discription
        reaction_disc = rdisc(i);
        disp('-----------------');
        
        if reaction_disc==1
            disp('Disscosiation to Gas-phase reaction : AB(g) --> A* + B*');
            % Calculation of Forward Activation energy
            E_AB(i,1) = 0.5*(D_AB+(Q_metal_atom_A*Q_metal_atom_B/(Q_metal_atom_A+Q_metal_atom_B))-Q_AB-Q_metal_atom_B);
            
            % if E_AB <0 then it should be equal zero
            if E_AB(i,1)<0    
                E_AB(i,1) = 0;
            end
            
            % Calculation of Reverse Activation energy
            E_AB(i,2) = E_AB(i,1) - dH;
            
            % if E_AB <0 then it should be equal zero
            if E_AB(i,2)<0    
               E_AB(i,2) = 0;
            end
            
        elseif reaction_disc==2
            disp('Recombination to Gas-phase reactions : A* + B* --> AB(g)')
            % Calculation of Forward Activation energy
            Eg = 0.5*(D_AB+(Q_metal_atom_A*Q_metal_atom_B/(Q_metal_atom_A+Q_metal_atom_B))-Q_AB-Q_metal_atom_B);
            E_AB(i,1)= Q_metal_atom_A + Q_metal_atom_B - D_AB + Eg;
            
            % if E_AB <0 then it should be equal zero
            if E_AB(i,1)<0    
                E_AB(i,1) = 0;
            end
            
            % Calculation of Reverse Activation energy
            E_AB(i,2) = E_AB(i,1) - dH;
            
            % if E_AB <0 then it should be equal zero
            if E_AB(i,2)<0    
               E_AB(i,2) = 0;
            end
            
        else
            disp('*** Error ***');
        end
        disp('-----------------');
        
    elseif reaction_type==2
        disp('Surface reactions    AB*    ')
        disp('Choose reactions discription : ');
        disp('Disscosiation an adsrorbate species  : AB* --> A* + B* = 1');
        disp('Recombination to chemisorbed         : A* + B* --> AB* = 2');

        % Determination of i th reaction discription
        reaction_disc = rdisc(i);
        disp('-----------------');
        
        if reaction_disc==1
            disp('Disscosiation an adsrorbate species  : AB* --> A* + B* ');
            
            % Calculation of Forward Activation energy
            E_AB(i,1) = 0.5*(D_AB+(Q_metal_atom_A*Q_metal_atom_B/(Q_metal_atom_A+Q_metal_atom_B))- Q_AB - Q_metal_atom_B)+ Q_AB;
            
            % if E_AB <0 then it should be equal zero
            if E_AB(i,1)<0    
                E_AB(i,1) = 0;
            end
            
            % Calculation of Reverse Activation energy
            E_AB(i,2) = E_AB(i,1) - dH;
            
            % if E_AB <0 then it should be equal zero
            if E_AB(i,2)<0    
               E_AB(i,2) = 0;
            end
            
        elseif reaction_disc==2
            disp('recombination to chemisorbed  :   A* + B* --> AB*');
            
            % Calculation of Forward Activation energy
            Eg = 0.5*(D_AB+(Q_metal_atom_A*Q_metal_atom_B/(Q_metal_atom_A+Q_metal_atom_B))- Q_AB - Q_metal_atom_B);

            % if E_AB <0 then it should be equal zero
            if Eg >0      
                E_AB(i,1)= Q_metal_atom_A + Q_metal_atom_B - D_AB + Eg;
            else
                E_AB(i,1)= Q_metal_atom_A + Q_metal_atom_B - D_AB;
            end

            % if E_AB <0 then it should be equal zero
            if E_AB(i,1)<0    
               E_AB(i,1) = 0;
            end
            
            % Calculation of Reverse Activation energy
            E_AB(i,2) = E_AB(i,1) - dH;
            
            % if E_AB <0 then it should be equal zero
            if E_AB(i,2)<0    
                E_AB(i,2) = 0;
            end

        else
            disp('*** Error ***');
        end
        disp('-----------------');
    elseif reaction_type==3
        disp('Choose reactions discription : ');
        disp('Chemisorption of a gasous species          : AB + * --> AB* = 1');
        disp('Production of a chemisorbed species        : AB* --> AB + * = 2');
        if reaction_disc==1
            disp('Chemisorption of a gasous species  :   AB + * --> AB*');
            E_AB(i,1) = 0;
            E_AB(i,2) = E_AB(i,1)-dH;
            
        elseif reaction_disc==2
            disp('Production of a chemisorbed species  :   AB* --> AB + *');
            dH = -dH;
            E_AB(i,2)= 0;
            E_AB(i,1)= E_AB(i,2)- dH;
        else
             disp('*** Error ***');
        end
    
    elseif reaction_type==4
        disp('Disproportation reactions X* + Y* --> Z* + F* ');
        
        % Q_metal_atom_A  and Q_metal_atom_B : The value of QA-M for this species
        % 1- Q_metal_atom_A  and Q_metal_atom_B are component so choose its heat of adsorbtion
        for t=1:component_num
            if No_A(i)==t
                Q_XYZF(i,3)=Q_species(t);
            end
            if No_B(i)==t
                Q_XYZF(i,4)=Q_species(t);
            end
        end

        % 2- Q_metal_atom_A  and Q_metal_atom_B are Atom so choose its heat of adsorbtion
        % 100 is just a symbol for counter (Example: for No. H=101=Q(1), C=102=Q(2), O=103=Q(3))
        for h=1:main_atom_num
            if No_A(i)==(h+100)
                Q_XYZF(i,3)=Q_metal_atom(h);
            end
            if No_B(i)==(h+100)
                Q_XYZF(i,4)=Q_metal_atom(h);
            end
        end 
        
        Q_XYZF(i,1)= Q_calculation_E(i,1);
        Q_XYZF(i,2)= Q_calculation_E(i,2);
        
        D_X = D_XYZF(i,1);
        D_Y = D_XYZF(i,2);
        D_Z = D_XYZF(i,3);
        D_F = D_XYZF(i,4);
        Q_X = Q_XYZF(i,1);
        Q_Y = Q_XYZF(i,2);
        Q_Z = Q_XYZF(i,3);
        Q_F = Q_XYZF(i,4);
        D_XY = D_X + D_Y - D_Z - D_F;
        Q_XY = Q_X + Q_Y;
        
        % Calculation of Forward Activation energy
        E_AB(i,1) = 0.5*(D_XY +((Q_Z*Q_F)/(Q_Z + Q_F)) + Q_XY - Q_Z - Q_F); 
        
        % Enthaply of i th reaction
        dH = D_XY + Q_XY - Q_Z - Q_F;
        
        % if E_AB <0 then it should be equal zero
        if E_AB(i,1)<0    
           E_AB(i,1) = 0;
        end
        
        % Calculation of Reverse Activation energy
        E_AB(i,2) = E_AB(i,1) - dH;
        
        % if E_AB <0 then it should be equal zero
        if E_AB(i,2)<0    
            E_AB(i,2) = 0;
        end
    else
        disp('*** Error ***');
    end
    disp('Enthalpy :  ')
    disp(dH)
    disp('-----------------');
    disp('Activation energy of i th species :  ')
    disp(E_AB(i,:))
end
disp('-----------------');
disp('-----------------');

% The matrix of Activation Energy
disp('Activation Energy ... E_forward and E_reverse :  ')
disp(E_AB)

% Build E matrix to enter in Arrinious Eq.
E=E_AB;
%%
% Arrinious Eq. and building the k (equation constant)
    for ii=1:main_reaction_num
        for jj=1:2
        k(ii,jj)= k0(ii,jj)*exp(-E(ii,jj)/(R*T));
        end
    end
    
    % Option to display output
options=optimset('Display','off'); 
    % Call optimizer
[x,fval] = fsolve(@myfun,x0,options);  
 
% The Kinetic Reaction equation (Teta, Folw rate, Pressure, Volume)
function F = myfun(x)
            F = [ k(1,1)*p(1)*x(4)*x(4)-k(1,2)*x(1)*x(1)-k(3,1)*x(2)*x(1)+k(3,2)*x(3)
                  k(2,1)*p(2)*x(4)-k(2,2)*x(2)-k(3,1)*x(2)*x(1)+k(3,2)*x(3)
                  k(3,1)*x(1)*x(2)-k(3,2)*x(3)-k(4,1)*x(3)+k(4,2)*p(3)*x(4)
                  x(1)+x(2)+x(3)+x(4)-1
                  flo(1)-x(5)-(k(1,1)*p(1)*x(4)*x(4)-k(1,2)*x(1)*x(1))*v
                  flo(2)-x(6)-(k(2,1)*p(2)*x(4)-k(2,2)*x(2))*v
                  flo(3)-x(7)+(k(4,1)*x(3)-k(4,2)*p(3)*x(4))*v];
end

        display(x);
        display(fval);
        display(xexp-x);
        
        %WRITE x and Q_metal_atom in matrix
        for d=1:7
            G(r,d)=x(d);
        end
        for d=1:3
            Q(r,d)=Q_metal_atom(d);
            QQQ(r,d)=Q_species(d);
        end
        % counter for fminsearch
        r=r+1;
        
        % Error function
        error=sum((xexp-x).^2);
    end
display(r)
xlswrite('testdata6.xlsx', G, 1,'A1:G700');
xlswrite('testdata6.xlsx', Q, 2,'A1:C700');
%xlswrite('testdata6.xlsx', QQ, 3,'A1:C600');
xlswrite('testdata6.xlsx', QQQ, 4,'A1:C700');
end
