select art_code from nse_dat where act_code ='RB' and nse_nums =''


select art.art_code, art.art_desl, nse.nse_nums from ART_PAR as art
left outer join nse_dat as nse on
nse.act_code = art.act_code and nse.art_code = art.art_code
and nse.act_code ='rb' --and nse.nse_nums =''

