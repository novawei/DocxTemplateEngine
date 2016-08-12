//
//  DocxTemplateEngine.m
//  DocxTemplateEngine
//
//  Created by WeiXinxing on 16/8/12.
//  Copyright © 2016年 nfs. All rights reserved.
//

#import "DocxTemplateEngine.h"
#import "SSZipArchive.h"
#import "MGTemplateEngine.h"
#import "ICUTemplateMatcher.h"

@interface DocxTemplateEngine () <MGTemplateEngineDelegate>

@property (nonatomic, strong) MGTemplateEngine *templateEngine;
@property (nonatomic, strong) NSString *workspaceDirectoryPath;
- (void)createWorkspaceIfNeeded;

@end

@implementation DocxTemplateEngine

+ (instancetype)sharedEngine
{
    static DocxTemplateEngine *instance;
    static dispatch_once_t onceToken;
    dispatch_once(&onceToken, ^{
        instance = [[DocxTemplateEngine alloc] init];
    });
    return instance;
}

- (instancetype)init
{
    if (self = [super init]) {
        self.workspaceDirectory = NSDocumentDirectory;
        self.workspaceName = @"DocxWorkspace";
        
        [self createWorkspaceIfNeeded];
        [self createTemplateEngine];
    }
    return self;
}

- (void)createTemplateEngine
{
    self.templateEngine = [[MGTemplateEngine alloc] init];
    self.templateEngine.matcher = [ICUTemplateMatcher matcherWithTemplateEngine:self.templateEngine];
    self.templateEngine.delegate = self;
}

- (void)createWorkspaceIfNeeded
{
    self.workspaceDirectoryPath = NSSearchPathForDirectoriesInDomains(self.workspaceDirectory, NSUserDomainMask, YES)[0];
    self.workspaceDirectoryPath = [self.workspaceDirectoryPath stringByAppendingPathComponent:self.workspaceName];
    if ([[NSFileManager defaultManager] fileExistsAtPath:self.workspaceDirectoryPath]) {
        return;
    }
    
    [[NSFileManager defaultManager] createDirectoryAtPath:self.workspaceDirectoryPath
                              withIntermediateDirectories:YES
                                               attributes:nil
                                                    error:nil];
}

- (NSString *)createDocxWithName:(NSString *)name 
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
{
    NSString *tplBundlePath = [[NSBundle mainBundle] pathForResource:tplname ofType:@"docx"];
    NSString *tplWorkspacePath = [self.workspaceDirectoryPath stringByAppendingFormat:@"/%@.docx", tplname];
    if ([[NSFileManager defaultManager] fileExistsAtPath:tplWorkspacePath]) {
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
    }
    [[NSFileManager defaultManager] copyItemAtPath:tplBundlePath toPath:tplWorkspacePath error:nil];
    // create tmp directory
    NSString *tmpDirectoryPath = [self.workspaceDirectoryPath stringByAppendingPathComponent:@"tmp"];
    if ([[NSFileManager defaultManager] fileExistsAtPath:tmpDirectoryPath]) {
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
    }
    [[NSFileManager defaultManager] createDirectoryAtPath:tmpDirectoryPath
                              withIntermediateDirectories:YES
                                               attributes:nil
                                                    error:nil];
    // unzip
    BOOL docxRet = [SSZipArchive unzipFileAtPath:tplWorkspacePath toDestination:tmpDirectoryPath];
    if (!docxRet) { // cleanup
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
        return nil;
    }
    
    // template replacing
    docxRet = NO;
    // word/document.xml
    NSString *contentPath = [tmpDirectoryPath stringByAppendingPathComponent:@"word/document.xml"];
    NSString *content = [self.templateEngine processTemplateInFileAtPath:contentPath withVariables:data];
    if (content.length > 0) {
        docxRet = [content writeToFile:contentPath atomically:YES encoding:NSUTF8StringEncoding error:nil];
    }
    if (!docxRet) { // cleanup
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
        return nil;
    }
    
    // zip
    NSString *zipPath = [self.workspaceDirectoryPath stringByAppendingFormat:@"/%@.zip", name];
    docxRet = [SSZipArchive createZipFileAtPath:zipPath withContentsOfDirectory:tmpDirectoryPath];
    if (!docxRet) { // cleanup
        [[NSFileManager defaultManager] removeItemAtPath:zipPath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
        return nil;
    }
    
    // rename
    NSString *docxPath = [self.workspaceDirectoryPath stringByAppendingFormat:@"/%@.docx", name];
    if ([[NSFileManager defaultManager] fileExistsAtPath:docxPath]) {
        [[NSFileManager defaultManager] removeItemAtPath:docxPath error:nil];
    }
    [[NSFileManager defaultManager] moveItemAtPath:zipPath toPath:docxPath error:nil];
    // cleanup
    [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
    [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
    
    return docxPath;
}

+ (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
{
    return [[DocxTemplateEngine sharedEngine] createDocxWithName:name
                                                            data:data
                                                   usingTemplate:tplname];
}

#pragma mark - 

- (void)templateEngine:(MGTemplateEngine *)engine encounteredError:(NSError *)error isContinuing:(BOOL)continuing
{
    NSLog(@"%@", error);
}

@end
