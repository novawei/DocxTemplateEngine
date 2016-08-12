//
//  DocxTemplateEngine.h
//  DocxTemplateEngine
//
//  Created by WeiXinxing on 16/8/12.
//  Copyright © 2016年 nfs. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface DocxTemplateEngine : NSObject

@property (nonatomic, assign) NSSearchPathDirectory workspaceDirectory; // NSDocumentDirectory
@property (nonatomic, strong) NSString *workspaceName; // DocxWorkspace
@property (nonatomic, strong) NSString *placeholderFormat; // {{%@}}

+ (instancetype)sharedEngine;

/**
 *  create docx (word document)
 *
 *  @param name    filename, e.g. name=@"test", the docx's filename will be "test.docx"
 *  @param data    dictionary with key and value, key MUST defined in config file
 *  @param tplname template name (in bundle), using default config named "name.json"
 *
 *  @return docx path if success, otherwise nil
 */
- (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname;
+ (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname;

/**
 *  create docx (word document)
 *
 *  @param name    filename, e.g. name=@"test", the docx's filename will be "test.docx"
 *  @param tplname template name (in bundle), using default config named "name.json"
 *  @param confname config file name (in bundle), json format
 *
 *  @return docx path if success, otherwise nil
 */
- (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
               andTemplateConfig:(NSString *)confname;
+ (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
               andTemplateConfig:(NSString *)confname;

@end
