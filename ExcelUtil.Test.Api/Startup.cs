using Magicodes.ExporterAndImporter.Excel;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.OpenApi.Models;
using System;
using System.IO;

namespace ExcelUtil.Test.Api
{
    /// <summary>
    ///
    /// </summary>
    public class Startup
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="configuration"></param>
        /// <param name="environment"></param>
        public Startup(IConfiguration configuration, IHostEnvironment environment)
        {
            Configuration = configuration;
            HostingEnvironment = environment;
        }

        /// <summary>
        ///
        /// </summary>
        public IHostEnvironment HostingEnvironment { get; }

        /// <summary>
        ///
        /// </summary>
        public IConfiguration Configuration { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="services"></param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();
            // 注册 HttpContext 上下文
            services.AddHttpContextAccessor();

            #region SwaggerGen

            services.AddSwaggerGen(options =>
            {
                options.SwaggerDoc("v1", new OpenApiInfo
                {
                    Version = "v1.0.0.0",
                    Title = "Excel导入测试项目Api",
                    Description = "Excel导入测试项目Api，通过调用本地封装库进行相关接口测试，针对导入模块。",
                });

                options.IncludeXmlComments(Path.Combine(AppContext.BaseDirectory, "ExcelUtil.Test.Api.xml"));
            });

            #endregion SwaggerGen

            #region 解决循环引用

            //FLAG:方案2：延迟加载
            services.AddMvc().AddNewtonsoftJson(opts =>
            {
                opts.SerializerSettings.ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore;
            });

            #endregion 解决循环引用

            //注册后台普通服务
            services.AddExcelService();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="app"></param>
        /// <param name="env"></param>
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            // 挂载 Swagger
            if (env.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI(options =>
                {
                    options.SwaggerEndpoint("/swagger/v1/swagger.json", "ExcelUtil.Test.Api V1");
                    options.RoutePrefix = string.Empty;
                });
            }

            app.UseRouting();

            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}