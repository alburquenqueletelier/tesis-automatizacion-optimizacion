  #R: v3.6.1
  #OS: Windows 10-64bits
  #Funcion programcion estocastica #
  #Copia modificada del algoritmo original "Stochastic_Programming_2S_GARMA_Example" #
  #el cual convierte el script orginal en una funcion con los siguientes parametros: #
  # datapro = demanda mensual de un producto, arreglo
  # precio = precio promedio de un producto, numerico
  # tope = presupuesto de un producto, numerico
  # ajuste = 'garma' o 'gamlss'
  # k_centros = numero de cluster a usar, numerico
  # q = cuartil a usar
  estocastica <- function(datapro,precio,tope,O,S,H,ajuste='garma',k_centros=10, q){
    library(gamlss.util)
    library(RcmdrMisc)
    library(e1071)
    library(glarma)
    #library(Rcmdr)
    library(forecast)
    
     source('ls_ists.R')

        Dataset=datapro
        o = O # Costo de ordenar mensual $/orden 
        s = S # Costo de desabastecimiento de un producto, Cantidad en dinero total comprada/Cantidad total de productos   (pendiente) 
        h = H # Costo de almacenar un producto por año  
        n = dim(Dataset)[1]-1 # segun tu cantidad de meses
        u = precio # Costo unitario del producto promedio (Varia por producto) 
        C = tope # Tope de presupuesto para compra, maxima compra que exista para el producto (varia por producto) 
        y=matrix(Dataset,ncol=1) # Cambiar el producto
        X=matrix(y,ncol=1)
        iteracion = 0
        sigma=as.numeric(histDist(y,family="NBII")$sigma)
        if (ajuste=='garma'){
            glm_model = garmaFit(y~X, order=c(1,0), family=NBII, tail=3) #Modelo GARMA con AR=1
            intercept=  as.numeric(glm_model$mu.coefficient)[1] #intercepto
            reg= as.numeric(glm_model$mu.coefficient)[2] #beta covariable
            phi=  as.numeric(glm_model$mu.coefficient)[3] # coeficiente autoregresivo
            theta= as.numeric(glm_model$mu.coefficient)[4] # Coeficiente media movil
            #################################ESTA PARTE DEBE SER MEJORADA######################################   
            mu = intercept+reg*X[length(y)]+phi*(fitted(glm_model)[length(y)-1]-intercept-reg*X[length(y)-1])
            #################################ESTA PARTE DEBE SER MEJORADA###################################### 
            dev.off()
        }else if (ajuste=='gamlss'){
            mu=as.numeric(fitted(gamlss(y~1,type="NBII"))[length(y)-1]) ### en caso de error hessian
        }else {
            stop('Se ingreso un modelo incorrecto')
        }
        #################################ESTA PARTE DEBE SER MEJORADA####################################
    
    
         family = "NBII"
        BB = 10000
         T = 1
         #q = 50  # como parametro 
     
     
     
    ###############################################################
    
    NREP = 1
    
    ###############################################################
    
       TC_glm_q1 = numeric()
       
       for(j in 1:NREP){
          
          
    glm_fit0 = mu    
    
    s_glm   = sigma
    
    glm_fit = glm_fit0
    
    if (glm_fit <= 0){
    stop ('Mu <= 0 en qNBII')
    }
    
    q1_glm   = qNBII(  (q/100),   glm_fit, sigma = s_glm)
    
    Yglmq1 = matrix(numeric(), ncol = T, nrow = BB)
    
     for(i in 1:T){
     ## se aplica esta condicional ya que no se logro 
     ## encontrar el motivo por el cual se genera el 
     ## error MU must be greater than 0
     if (q1_glm[i] < 0){
         q1_glm[i] = abs(q1_glm[i])
     }
     
     if (q1_glm[i] <= 0){
     stop('Mu <= 0 en rNBII')
     }
     y_aux_glmq1   = rNBII(BB, q1_glm[i], s_glm) 
     
     Yglmq1[,i]   = y_aux_glmq1
     
          }
     
     apply(Yglmq1, 2, mean)
     apply(Yglmq1, 2, sd)
     
     
    ################################################################
    ####### OBTAIN CENTROIDS AND PROBABILITY BY CLUSTER ANALYSIS  
    ################################################################
    
    pr_glmq1 = 
    matrix(numeric(), nrow = 10, ncol = T) 
    
    center_glmq1 = 
    matrix(numeric(), nrow = 10, ncol = T) 
    
     for(t in 1:T){
     km_glmq1   = KMeans(model.matrix(~-1 +   Yglmq1[,t]), centers = k_centros,
           iter.max = 10, num.seeds = 12)
           
     pr_glmq1[,t]   = c(table(km_glmq1$cluster)  /sum(table(km_glmq1$cluster))   ,
                     rep(NA, 10 - 10))
     
     center_glmq1[,t]   = c(as.numeric(km_glmq1$center)  , rep(NA, 10 - 10))
                                                                               
     }
    
     
    ################################################################
    ####### OPTIMIZE MODEL FOR GLM Q1
    ################################################################
    
    
    ########## Optimization
    opt1_glmq1 = stsslp(          # Shortage of two-stage stochastic linear programming:
    y = as.numeric(na.omit(center_glmq1[,1])), # centers KMeans Cluster DPUT to be satisfied in period.
    p = as.numeric(na.omit(pr_glmq1[,1])),     # probability of scenario.
    o,                      # Fixed order cost in period.
    h,                      # Holding cost at the end of period.
    s,                      # Shortage cost at the end of period.
    C,                      # Budget purchase in period.
    u)                      # Unitary cost of purchase in period.
    
    
    
    Q1_glmq1 = opt1_glmq1$Inv_mod_results[2]       # Cantidad a pedir
    
    Z1_glmq1 = opt1_glmq1$Inv_mod_results[1]       # V.Binaria ordenar
    
    I1_glmq1 = opt1_glmq1$Inv_mod_results[3]       # Saldos
    
    B1_glmq1 = opt1_glmq1$Inv_mod_results[4]       # Desabastecidos
     
          }
          
    
   # opt1_glmq1
    
    s=sum(pr_glmq1*matrix(opt1_glmq1$B,ncol=1))
    resultado <- c(opt1_glmq1,s,mu,sigma,q1_glm) # cuando volver a pedir
    rm(list=setdiff(ls(), "resultado"))
    resultado
    }
    
