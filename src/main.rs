use rust_xlsxwriter::*;
// use calamine::*;


fn main() -> Result<(), Box<dyn std::error::Error>>{
    let header: [&'static str; 195]=[
        // General
        "CHROM","POS","REF","ALT","DP","AD","QUAL","MQ","Zygosity","FILTER","Effect",
        "Putative_Impact","Gene_Name","Feature_Type","Feature_ID","Transcript_BioType",
        "Rank/Total","HGVS.c","HGVS.p","REF_AA","ALT_AA","cDNA_pos","cDNA_length","CDS_pos",
        "CDS_length","AA_pos","AA_length","Distance",
        // dbSNP138 annotation
        "dbSNP138_ID","dbSNP156_ID",
        // 1000 genomes phase 3 annotation
        "p3_1000G_AF","p3_1000G_AFR_AF","p3_1000G_AMR_AF","p3_1000G_EAS_AF","p3_1000G_EUR_AF","p3_1000G_SAS_AF",
        // EVS annotation
        "ESP6500_MAF_EA","ESP6500_MAF_AA","ESP6500_MAF_ALL",
        // Clinvar annotation
        "CLINVAR_CLNSIG","CLINVAR_CLNDISDB","CLINVAR_CLNDN","CLINVAR_CLNREVSTAT",
        "ACMG_SF_v3.2","REF_AA_dbnsfp","ALT_AA_dbnsfp","hg38_chr","hg38_pos(1-based)","cds_strand","refcodon","codonpos",
        "codon_degeneracy","SIFT_score","SIFT_converted_rankscore","SIFT_pred","LRT_score",
        "LRT_converted_rankscore","LRT_pred","LRT_Omega","MutationTaster_score",
        "MutationTaster_converted_rankscore","MutationTaster_pred","MutationTaster_model",
        "MutationTaster_AAE","MutationAssessor_score","MutationAssessor_rankscore",
        "MutationAssessor_pred","FATHMM_score","FATHMM_converted_rankscore","FATHMM_pred",
        "PROVEAN_score","PROVEAN_converted_rankscore","PROVEAN_pred","MetaSVM_score",
        "MetaSVM_rankscore","MetaSVM_pred","MetaLR_score","MetaLR_rankscore","MetaLR_pred",
        "Reliability_index","M-CAP_score","M-CAP_rankscore","M-CAP_pred","MutPred_score",
        "MutPred_rankscore","MutPred_protID","MutPred_AAchange","MutPred_Top5features",
        "fathmm-MKL_coding_score","fathmm-MKL_coding_rankscore","fathmm-MKL_coding_pred",
        "fathmm-MKL_coding_group","Eigen-raw_coding","Eigen-phred_coding",
        "Eigen-PC-raw_coding","Eigen-PC-phred_coding","Eigen-PC-raw_coding_rankscore",
        "integrated_fitCons_score","integrated_fitCons_rankscore","integrated_confidence_value",
        "GERP++_NR","GERP++_RS","GERP++_RS_rankscore","gnomAD_exomes_AC","gnomAD_exomes_AN",
        "gnomAD_exomes_AF","gnomAD_exomes_AFR_AC","gnomAD_exomes_AFR_AN","gnomAD_exomes_AFR_AF",
        "gnomAD_exomes_AMR_AC","gnomAD_exomes_AMR_AN","gnomAD_exomes_AMR_AF",
        "gnomAD_exomes_ASJ_AC","gnomAD_exomes_ASJ_AN","gnomAD_exomes_ASJ_AF",
        "gnomAD_exomes_EAS_AC","gnomAD_exomes_EAS_AN","gnomAD_exomes_EAS_AF",
        "gnomAD_exomes_FIN_AC","gnomAD_exomes_FIN_AN","gnomAD_exomes_FIN_AF",
        "gnomAD_exomes_NFE_AC","gnomAD_exomes_NFE_AN","gnomAD_exomes_NFE_AF",
        "gnomAD_exomes_SAS_AC","gnomAD_exomes_SAS_AN","gnomAD_exomes_SAS_AF",
        "gnomAD_genomes_AC","gnomAD_genomes_AN","gnomAD_genomes_AF","gnomAD_genomes_AFR_AC",
        "gnomAD_genomes_AFR_AN","gnomAD_genomes_AFR_AF","gnomAD_genomes_AMR_AC",
        "gnomAD_genomes_AMR_AN","gnomAD_genomes_AMR_AF","gnomAD_genomes_ASJ_AC",
        "gnomAD_genomes_ASJ_AN","gnomAD_genomes_ASJ_AF","gnomAD_genomes_EAS_AC",
        "gnomAD_genomes_EAS_AN","gnomAD_genomes_EAS_AF","gnomAD_genomes_FIN_AC",
        "gnomAD_genomes_FIN_AN","gnomAD_genomes_FIN_AF","gnomAD_genomes_NFE_AC",
        "gnomAD_genomes_NFE_AN","gnomAD_genomes_NFE_AF","Interpro_domain","GTEx_V8_gene",
        "GTEx_V8_tissue","MIM_id","Gene_old_names","Gene_full_name","Pathway(Uniprot)",
        "Pathway(BioCarta)_short","Pathway(BioCarta)_full","Pathway(ConsensusPathDB)",
        "Pathway(KEGG)_id","Pathway(KEGG)_full","Function_description","Disease_description",
        "MIM_phenotype_id","MIM_disease","Trait_association(GWAS)","GO_biological_process",
        "GO_cellular_component","GO_molecular_function","Tissue_specificity(Uniprot)",
        "Expression(egenetics)","Expression(GNF/Atlas)","Interactions(IntAct)",
        "Interactions(BioGRID)","Interactions(ConsensusPathDB)","P(HI)","P(rec)",
        "Known_rec_info","RVIS_EVS","RVIS_percentile_EVS","LoF-FDR_ExAC","RVIS_ExAC",
        "RVIS_percentile_ExAC","GHIS","GDI","GDI-Phred",
        "Gene_damage_prediction(all_disease-causing_genes)",
        "Gene_damage_prediction(all_Mendelian_disease-causing_genes)",
        "Gene_damage_prediction(Mendelian_AD_disease-causing_genes)",
        "Gene_damage_prediction(Mendelian_AR_disease-causing_genes)",
        "Gene_damage_prediction(all_PID_disease-causing_genes)",
        "Gene_damage_prediction(PID_AD_disease-causing_genes)",
        "Gene_damage_prediction(PID_AR_disease-causing_genes)",
        "Gene_damage_prediction(all_cancer_disease-causing_genes)",
        "Gene_damage_prediction(cancer_recessive_disease-causing_genes)",
        "Gene_damage_prediction(cancer_dominant_disease-causing_genes)"
    ];
    
    let mut result_workbook = Workbook::new();
    let format = Format::new()
        .set_bold()
        .set_font_size(10)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_background_color(Color::RGB(0xADCAE6))
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::Black)
        .set_text_wrap();
    let result_worksheet = result_workbook.add_worksheet();

    for (i, &col_name) in header.iter().enumerate() {
        result_worksheet.write_with_format(0, i as u16, col_name, &format)?;
    }

    result_workbook.save("variant_annotation.xlsx")?;
    Ok(())
}
