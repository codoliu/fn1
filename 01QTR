        public JsonResult UploadProcfile(IEnumerable<HttpPostedFileBase> files, string itemId, string statusId)
        {
            string msg = "Failed!";
            if (files != null)
            {
                foreach (var file in files)
                {
                    // Verify that the user selected a file
                    if (file != null && file.ContentLength > 0)
                    {
                        string fileName = System.IO.Path.GetFileName(file.FileName.Replace("/", "").Replace(@"\", "").Replace(@":", "").Replace(@"*", "").Replace(@"?", "").Replace(@"""", "").Replace(@"<", "").Replace(@">", ""));
                        //string path = System.IO.Path.Combine(Server.MapPath("~/"), "_RptOutput/" + fileName);
                        string extension = Path.GetExtension(file.FileName.Replace("/", "").Replace(@"\", ""));
                        //file.SaveAs(path);
                        try
                        {
                            bool convertSuccess = false;
                            if (extension.ToUpper().Equals(".XLS"))
                            {
                                fileName = file.FileName + "x";
                                string path = System.IO.Path.Combine(Server.MapPath("~/"), @"_RptOutput\" + fileName.Replace("/", "").Replace(@"\", ""));
                                extension = ".xlsx";
                                Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
                                workbook.LoadFromStream(file.InputStream);
                                workbook.SaveToFile(path, Spire.Xls.ExcelVersion.Version2010);
                                convertSuccess = true;
                            }
                            else if (extension.ToUpper().Equals(".XLSX"))
                            {
                                string path = System.IO.Path.Combine(Server.MapPath("~/"), @"_RptOutput\" + fileName.Replace("/", "").Replace(@"\", ""));
                                file.SaveAs(path);
                                convertSuccess = true;
                                //string path = System.IO.Path.Combine(Server.MapPath("~/"), @"_RptOutput\" + fileName);
                                //Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
                                //workbook.LoadFromStream(file.InputStream);
                                //workbook.SaveToFile(path, Spire.Xls.ExcelVersion.Version2010);
                                //convertSuccess = true;
                            }
                            else
                            {
                                return Json("上傳檔案非EXCEL!");
                            }

                            if (convertSuccess)
                            {
                                ProcfilesModel fileModel = new ProcfilesModel();
                                fileModel.id = Guid.NewGuid();
                                fileModel.exptapplyitemid = model.itemId; // Guid.Parse(itemId);
                                fileModel.itemstatus = model.statusId; // Guid.Parse(statusId);
                                fileModel.filename = fileName;
                                expt.ExecuteProcfiles(Service.Expt.EditMode.Insert, new List<ProcfilesModel> { fileModel });

                                msg = "Success!";
                            }
                        }
                        catch (Exception ex)
                        {
                            msg = ex.Message;
                        }
                    }
                }
            }
            return Json(msg);
        }
